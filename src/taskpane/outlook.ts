/* global document, Office, console, window, HTMLButtonElement, FormData, Blob, Response, fetch */

import { authenticateWithNaa, formatAuthenticationError, isNestedAppAuthSupported } from "./auth";

type MailAttachment = Office.AttachmentDetails & { contentType?: string };
type ApiResponse = Record<string, unknown>;
type JwtPayload = Record<string, unknown>;

const ANALYZE_ENDPOINT =
  "https://sistematdr-g4c5asddesgbd7bh.southcentralus-01.azurewebsites.net/api/tdr/analizar";

let compatibleAttachments: MailAttachment[] = [];
let retrievalId = 0;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  document.getElementById("sideload-msg")?.classList.add("is-hidden");
  document.getElementById("app-body")?.classList.remove("is-hidden");
  document.getElementById("refresh")?.addEventListener("click", renderCurrentMessage);
  document.getElementById("analyze")?.addEventListener("click", analyzeCompatibleAttachments);
  document.getElementById("test-auth")?.addEventListener("click", testAuthentication);
  updateAuthenticationAvailability();
  renderCurrentMessage();

  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, renderCurrentMessage);
  }
});

async function testAuthentication(): Promise<void> {
  const button = document.getElementById("test-auth") as HTMLButtonElement | null;
  if (!button) return;

  button.disabled = true;
  button.setAttribute("aria-busy", "true");
  showAuthenticationStatus("Autenticando con Microsoft...", "loading");

  try {
    const result = await authenticateWithNaa();
    const payload = decodeJwtPayload(result.accessToken);
    const diagnosticClaims = ["aud", "iss", "scp", "ver", "exp"]
      .map((claim) => `${claim}: ${formatJwtClaim(payload[claim])}`)
      .join("\n");
    showAuthenticationStatus(diagnosticClaims, "success");
  } catch (error: unknown) {
    const diagnostic = formatAuthenticationError(error);
    showAuthenticationStatus(`Error de autenticación: ${diagnostic}`, "error");
    console.error("Error de autenticación NAA/MSAL:", diagnostic);
  } finally {
    button.disabled = !isNestedAppAuthSupported();
    button.setAttribute("aria-busy", "false");
  }
}

function decodeJwtPayload(accessToken: string): JwtPayload {
  const segments = accessToken.split(".");
  if (segments.length !== 3 || !segments[1]) {
    throw new Error("El access token recibido no tiene un payload JWT válido.");
  }

  const base64 = segments[1].replace(/-/g, "+").replace(/_/g, "/");
  const paddedBase64 = base64.padEnd(Math.ceil(base64.length / 4) * 4, "=");
  const bytes = Uint8Array.from(window.atob(paddedBase64), (character) => character.charCodeAt(0));
  const encodedPayload = Array.from(bytes, (byte) => `%${byte.toString(16).padStart(2, "0")}`).join(
    ""
  );
  const payload: unknown = JSON.parse(decodeURIComponent(encodedPayload));

  if (typeof payload !== "object" || payload === null || Array.isArray(payload)) {
    throw new Error("El payload del access token no es un objeto JSON válido.");
  }

  return payload as JwtPayload;
}

function formatJwtClaim(value: unknown): string {
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  if (Array.isArray(value)) {
    return value.map(String).join(" ");
  }

  return "(ausente)";
}

function updateAuthenticationAvailability(): void {
  const button = document.getElementById("test-auth") as HTMLButtonElement | null;
  if (!button) return;

  const supported = isNestedAppAuthSupported();
  button.disabled = !supported;
  if (!supported) {
    showAuthenticationStatus(
      "NestedAppAuth 1.1 no está disponible en esta versión o cuenta de Outlook.",
      "error"
    );
  }
}

function showAuthenticationStatus(
  message: string,
  state: "idle" | "loading" | "success" | "error"
): void {
  const status = document.getElementById("auth-status");
  if (!status) return;
  status.textContent = message;
  status.className = `analysis-status analysis-status--${state}`;
  status.classList.toggle("is-hidden", state === "idle");
}

function renderCurrentMessage(): void {
  retrievalId += 1;
  const item = Office.context.mailbox.item;
  if (!item) {
    showStatus("Abre un correo para ver su información.", false);
    clearMessage();
    return;
  }

  const sender = item.from;
  const attachments = (item.attachments || []) as MailAttachment[];
  compatibleAttachments = attachments.filter(isCompatibleAttachment);

  setText("message-subject", item.subject?.trim() || "(Sin asunto)");
  setText("sender-name", sender?.displayName?.trim() || "Remitente desconocido");
  setText("sender-email", sender?.emailAddress?.trim() || "");
  document.getElementById("sender-email")?.classList.toggle("is-hidden", !sender?.emailAddress);
  renderCompatibleAttachments(compatibleAttachments);
  updateAnalyzeButton();
  if (compatibleAttachments.length === 0) {
    showAnalysisStatus("No se encontraron documentos PDF, DOCX o DOC para analizar.", "error");
  } else {
    showAnalysisStatus("", "idle");
  }
  showStatus("", true);
}

async function analyzeCompatibleAttachments(): Promise<void> {
  const item = Office.context.mailbox.item;
  const attachments = [...compatibleAttachments];
  if (!item) return;
  if (attachments.length === 0) {
    showAnalysisStatus("No se encontraron documentos PDF, DOCX o DOC para analizar.", "error");
    return;
  }
  const operationId = ++retrievalId;

  setAnalyzeBusy(true);
  showAnalysisStatus("Procesando TDR...", "loading");

  const internetMessageId = item.internetMessageId?.trim();
  if (!internetMessageId) {
    showAnalysisStatus(
      "El correo aún no tiene internetMessageId. Actualiza el correo e inténtalo nuevamente.",
      "error"
    );
    setAnalyzeBusy(false);
    return;
  }

  if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
    showAnalysisStatus(
      "Esta versión de Outlook no permite obtener el contenido del adjunto.",
      "error"
    );
    console.error("Mailbox 1.8 no está disponible; getAttachmentContentAsync no es compatible.");
    setAnalyzeBusy(false);
    return;
  }

  try {
    const attachmentContents = await Promise.all(
      attachments.map(async (attachment) => ({
        attachment,
        content: await getAttachmentContent(item, attachment),
      }))
    );
    if (operationId !== retrievalId) return;

    for (const { attachment, content } of attachmentContents) {
      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        throw new Error(
          `El documento "${attachment.name}" tiene un formato de contenido no compatible: ${content.format}.`
        );
      }
    }

    const authentication = await authenticateWithNaa();
    if (operationId !== retrievalId) return;

    const formData = new FormData();
    formData.append("internetMessageId", internetMessageId);
    for (const { attachment, content } of attachmentContents) {
      const fileName = attachment.name || "documento";
      formData.append(
        "archivo",
        new Blob([decodeBase64(content.content)], { type: getAttachmentMimeType(fileName) }),
        fileName
      );
    }

    let response: Response;
    try {
      response = await fetch(ANALYZE_ENDPOINT, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${authentication.accessToken}`,
        },
        body: formData,
      });
    } catch (error: unknown) {
      throw new BackendNetworkError(error);
    }

    if (operationId !== retrievalId) return;
    const parsedResponse = await readJsonResponse(response);
    if (operationId !== retrievalId) return;
    showBackendResponse(response, parsedResponse);
  } catch (error: unknown) {
    if (operationId !== retrievalId) return;
    showAnalysisError(error);
  } finally {
    if (operationId === retrievalId) setAnalyzeBusy(false);
  }
}

function getAttachmentContent(
  item: Office.MessageRead,
  attachment: MailAttachment
): Promise<Office.AttachmentContent> {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachment.id, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(
          new Error(`No se pudo obtener el documento "${attachment.name}": ${result.error.message}`)
        );
        return;
      }

      resolve(result.value);
    });
  });
}

async function readJsonResponse(
  response: Response
): Promise<{ data?: ApiResponse; isJson: boolean }> {
  const text = await response.text();
  if (!text.trim()) return { isJson: false };

  try {
    const parsed: unknown = JSON.parse(text);
    if (typeof parsed === "object" && parsed !== null && !Array.isArray(parsed)) {
      return { data: parsed as ApiResponse, isJson: true };
    }
    return { isJson: false };
  } catch {
    return { isJson: false };
  }
}

function showBackendResponse(
  response: Response,
  parsed: { data?: ApiResponse; isJson: boolean }
): void {
  const data = parsed.data;
  if (!parsed.isJson || !data) {
    showAnalysisStatus(
      `El backend devolvió una respuesta no JSON (HTTP ${response.status}).`,
      "error"
    );
    return;
  }

  const status = typeof data.status === "string" ? data.status.toLocaleLowerCase() : "";
  const details = getApiMessage(data);

  if (response.status === 202) {
    const message =
      status === "processing"
        ? "El TDR continúa en procesamiento."
        : "La solicitud fue aceptada y continúa en procesamiento.";
    showAnalysisStatus(details || message, "loading");
    return;
  }

  if (response.status === 400) {
    showAnalysisStatus(
      details ? `Solicitud no válida: ${details}` : "Solicitud no válida (HTTP 400).",
      "error"
    );
    return;
  }

  if (response.status === 401 || response.status === 403) {
    showAnalysisStatus(
      details ||
        `No autorizado por el backend (HTTP ${response.status}). Revisa el token y los permisos.`,
      "error"
    );
    return;
  }

  if (!response.ok) {
    showAnalysisStatus(details || `El backend devolvió un error HTTP ${response.status}.`, "error");
    return;
  }

  if (data.alreadyProcessed === true && status === "completed") {
    showAnalysisStatus(details || "Este correo ya había sido procesado correctamente.", "success");
    return;
  }

  showAnalysisStatus(details || "TDR procesado correctamente.", "success");
}

function getApiMessage(data: ApiResponse): string | undefined {
  for (const key of ["message", "detail", "title", "error"]) {
    if (typeof data[key] === "string" && data[key].trim()) return data[key].trim();
  }
  return undefined;
}

class BackendNetworkError extends Error {
  readonly cause: unknown;

  constructor(cause: unknown) {
    super("No se pudo conectar con el backend. Revisa la conexión y la configuración CORS.");
    this.name = "BackendNetworkError";
    this.cause = cause;
  }
}

function showAnalysisError(error: unknown): void {
  if (error instanceof BackendNetworkError) {
    showAnalysisStatus(error.message, "error");
    console.error("Error de red/CORS al invocar el backend:", error.cause);
    return;
  }

  const diagnostic = formatAuthenticationError(error);
  showAnalysisStatus(`No se pudo procesar el TDR: ${diagnostic}`, "error");
  console.error("Error durante el análisis del TDR:", diagnostic);
}

function decodeBase64(content: string): Uint8Array<ArrayBuffer> {
  const binary = window.atob(content.replace(/\s/g, ""));
  const bytes = new Uint8Array(new ArrayBuffer(binary.length));
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function getAttachmentExtension(fileName: string): string {
  const extensionStart = fileName.lastIndexOf(".");
  return extensionStart >= 0 ? fileName.slice(extensionStart).toLocaleLowerCase() : "";
}

function isCompatibleAttachment(attachment: MailAttachment): boolean {
  return [".pdf", ".docx", ".doc"].includes(getAttachmentExtension(attachment.name || ""));
}

function getAttachmentMimeType(fileName: string): string {
  switch (getAttachmentExtension(fileName)) {
    case ".pdf":
      return "application/pdf";
    case ".docx":
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    case ".doc":
      return "application/msword";
    default:
      return "application/octet-stream";
  }
}

function renderCompatibleAttachments(attachments: MailAttachment[]): void {
  const list = document.getElementById("document-list");
  const empty = document.getElementById("document-empty");
  const count = document.getElementById("document-count");
  if (!list || !empty || !count) return;

  list.replaceChildren();
  count.textContent = String(attachments.length);
  empty.classList.toggle("is-hidden", attachments.length > 0);

  attachments.forEach((attachment) => {
    const row = document.createElement("li");
    row.className = "attachment";
    const icon = document.createElement("span");
    icon.className = "attachment__icon";
    icon.setAttribute("aria-hidden", "true");
    icon.textContent = getAttachmentExtension(attachment.name || "")
      .slice(1)
      .toLocaleUpperCase();
    const details = document.createElement("div");
    details.className = "attachment__details";
    const name = document.createElement("span");
    name.className = "attachment__name";
    name.textContent = attachment.name || "Documento";
    const size = document.createElement("span");
    size.className = "attachment__size";
    size.textContent = formatFileSize(attachment.size);
    details.append(name, size);
    row.append(icon, details);
    list.appendChild(row);
  });
}

function formatFileSize(bytes?: number): string {
  if (!bytes || bytes < 1) return "Tamaño no disponible";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function showStatus(message: string, hidden: boolean): void {
  const status = document.getElementById("status");
  if (!status) return;
  status.textContent = message;
  status.classList.toggle("is-hidden", hidden);
}

function updateAnalyzeButton(): void {
  const button = document.getElementById("analyze") as HTMLButtonElement | null;
  if (button) button.disabled = compatibleAttachments.length === 0;
}

function setAnalyzeBusy(busy: boolean): void {
  const button = document.getElementById("analyze") as HTMLButtonElement | null;
  if (!button) return;
  button.disabled = busy || compatibleAttachments.length === 0;
  button.setAttribute("aria-busy", String(busy));
}

function showAnalysisStatus(
  message: string,
  state: "idle" | "loading" | "success" | "error"
): void {
  const status = document.getElementById("analysis-status");
  if (!status) return;
  status.textContent = message;
  status.className = `analysis-status analysis-status--${state}`;
  status.classList.toggle("is-hidden", state === "idle");
}

function clearMessage(): void {
  compatibleAttachments = [];
  setText("message-subject", "—");
  setText("sender-name", "—");
  setText("sender-email", "");
  renderCompatibleAttachments([]);
  updateAnalyzeButton();
  showAnalysisStatus("", "idle");
}

function setText(id: string, value: string): void {
  const element = document.getElementById(id);
  if (element) element.textContent = value;
}

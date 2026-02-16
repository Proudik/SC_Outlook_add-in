import { listAttachments, getAttachmentBase64 } from "./outlookAttachments";
import { uploadDocumentToCase } from "./singlecaseDocuments";

export async function uploadCurrentEmailAttachmentsToCase(params: {
  caseId: string;
}): Promise<{
  uploaded: Array<{ attachmentName: string; documentId: string }>;
  failed: Array<{ attachmentName: string; error: string }>;
}> {
  const { caseId } = params;

  const attachments = listAttachments();

  const uploaded: Array<{ attachmentName: string; documentId: string }> = [];
  const failed: Array<{ attachmentName: string; error: string }> = [];

  for (const att of attachments) {
    try {
      const content = await getAttachmentBase64(att);

      const resp = await uploadDocumentToCase({
        caseId,
        fileName: content.name,
        mimeType: content.mimeType,
        dataBase64: content.base64,
      });

      const first = resp.documents?.[0];
      uploaded.push({ attachmentName: att.name, documentId: first?.id ? String(first.id) : "" });
    } catch (e) {
      failed.push({
        attachmentName: att.name,
        error: e instanceof Error ? e.message : String(e),
      });
    }
  }

  return { uploaded, failed };
}

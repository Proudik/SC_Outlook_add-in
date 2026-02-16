type OutlookAttachment = {
    id: string;
    name: string;
    contentType?: string;
    isInline?: boolean;
  };
  
  export function listAttachments(): OutlookAttachment[] {
    const item = Office.context.mailbox.item;
    if (!item) throw new Error("No Outlook item.");
  
    return (item.attachments || []).filter((a) => !a.isInline) as OutlookAttachment[];
  }
  
  export function getAttachmentBase64(att: OutlookAttachment): Promise<{
    name: string;
    mimeType: string;
    base64: string;
  }> {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.getAttachmentContentAsync(
        att.id,
        (result: Office.AsyncResult<Office.AttachmentContent>) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(result.error?.message || "Failed to read attachment."));
            return;
          }
  
          const content = result.value;
          if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
            reject(new Error(`Unsupported attachment format: ${content.format}`));
            return;
          }
  
          resolve({
            name: att.name,
            mimeType: att.contentType || "application/octet-stream",
            base64: content.content,
          });
        }
      );
    });
  }
  
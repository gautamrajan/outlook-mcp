const config = require('../config');
const { callGraphAPIPaginated } = require('../utils/graph-api');

function mapAttachmentType(odataType) {
  switch (odataType) {
    case '#microsoft.graph.fileAttachment':
      return 'file';
    case '#microsoft.graph.itemAttachment':
      return 'item';
    case '#microsoft.graph.referenceAttachment':
      return 'reference';
    default:
      return 'unknown';
  }
}

function normalizeAttachment(graphAttachment) {
  const attachmentType = mapAttachmentType(graphAttachment['@odata.type']);
  const size = Number(graphAttachment.size) || 0;
  let downloadSupported = false;
  let downloadReason = '';

  if (attachmentType === 'file') {
    if (size <= config.MAX_ATTACHMENT_DOWNLOAD_BYTES) {
      downloadSupported = true;
    } else {
      downloadReason = 'attachment exceeds the maximum supported size of 25 MB';
    }
  } else if (attachmentType === 'item') {
    downloadReason = 'item attachments are not supported in v1';
  } else if (attachmentType === 'reference') {
    downloadReason = 'reference attachments are not supported in v1';
  } else {
    downloadReason = 'unknown attachment type';
  }

  return {
    id: graphAttachment.id,
    name: graphAttachment.name || 'Unnamed attachment',
    attachmentType,
    contentType: graphAttachment.contentType || 'application/octet-stream',
    size,
    isInline: graphAttachment.isInline === true,
    downloadSupported,
    downloadReason,
  };
}

async function listMessageAttachments(accessToken, emailId) {
  const endpoint = `me/messages/${encodeURIComponent(emailId)}/attachments`;
  const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint);
  const attachments = Array.isArray(response.value) ? response.value : [];
  return attachments.map(normalizeAttachment);
}

function formatAttachmentBlock(attachments) {
  return attachments.map((attachment, index) => {
    const supportText = attachment.downloadSupported
      ? 'Yes'
      : `No${attachment.downloadReason ? ` (${attachment.downloadReason})` : ''}`;

    return `${index + 1}. ${attachment.name}
ID: ${attachment.id}
Type: ${attachment.attachmentType}
Content Type: ${attachment.contentType}
Size: ${attachment.size} bytes
Inline: ${attachment.isInline ? 'Yes' : 'No'}
Download Supported: ${supportText}`;
  }).join('\n\n');
}

function formatAttachmentListText(emailId, attachments) {
  if (!attachments || attachments.length === 0) {
    return `No attachments found for email ${emailId}.`;
  }

  return `Found ${attachments.length} attachments for email ${emailId}:\n\n${formatAttachmentBlock(attachments)}`;
}

function formatAttachmentSummaryText(attachments) {
  if (!attachments || attachments.length === 0) {
    return 'Attachments:\nNone';
  }

  return `Attachments:\n${formatAttachmentBlock(attachments)}`;
}

function sanitizeAttachmentFilename(name) {
  const fallback = 'attachment.bin';
  const rawName = typeof name === 'string' && name.trim() ? name.trim() : fallback;
  return rawName
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
    .replace(/\s+/g, ' ')
    .slice(0, 255) || fallback;
}

module.exports = {
  listMessageAttachments,
  normalizeAttachment,
  formatAttachmentListText,
  formatAttachmentSummaryText,
  sanitizeAttachmentFilename,
};

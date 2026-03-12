const config = require('../config');
const { requestContext } = require('../auth/request-context');
const { streamGraphAPI } = require('../utils/graph-api');
const {
  lookupDownloadTicket,
  markDownloadTicketConsumed,
} = require('./download-ticket-store');

function sendText(res, statusCode, text, headers = {}) {
  if (typeof res.status === 'function') {
    res.status(statusCode);
    if (typeof res.set === 'function') {
      res.set(headers);
    }
    return res.send(text);
  }

  res.writeHead(statusCode, {
    'Content-Type': 'text/plain; charset=utf-8',
    ...headers,
  });
  res.end(text);
}

function disposeStream(stream) {
  if (!stream) return;
  if (typeof stream.resume === 'function') {
    stream.resume();
  }
  if (typeof stream.destroy === 'function') {
    stream.destroy();
  }
}

function encodeContentDispositionFilename(name) {
  return encodeURIComponent(name)
    .replace(/['()]/g, escape)
    .replace(/\*/g, '%2A');
}

async function handleAttachmentDownloadRequest(req, res, { token }) {
  const lookup = lookupDownloadTicket(token);

  if (lookup.status === 'invalid') {
    return sendText(res, 404, 'Attachment download URL is invalid.');
  }

  if (lookup.status === 'expired' || lookup.status === 'consumed') {
    return sendText(res, 410, 'Download URL expired. Request a new attachment download URL.');
  }

  const ticket = lookup.ticket;
  const endpoint = `me/messages/${encodeURIComponent(ticket.emailId)}/attachments/${encodeURIComponent(ticket.attachmentId)}/$value`;

  try {
    const upstream = await requestContext.run(ticket.authContext || null, async () => {
      return streamGraphAPI(ticket.accessToken, endpoint);
    });

    if (upstream.statusCode === 404) {
      disposeStream(upstream.stream);
      return sendText(res, 404, 'Attachment not found or no longer available.');
    }

    if (upstream.statusCode === 401 || upstream.statusCode === 403) {
      disposeStream(upstream.stream);
      return sendText(res, 410, 'Download URL expired. Request a new attachment download URL.');
    }

    if (upstream.statusCode < 200 || upstream.statusCode >= 300) {
      disposeStream(upstream.stream);
      return sendText(res, 502, 'Failed to download attachment from Microsoft Graph.');
    }

    const contentLength = Number(upstream.headers['content-length'] || ticket.size || 0);
    if (contentLength > config.MAX_ATTACHMENT_DOWNLOAD_BYTES) {
      disposeStream(upstream.stream);
      return sendText(res, 413, 'Attachment exceeds the maximum supported size of 25 MB.');
    }

    if (!markDownloadTicketConsumed(token)) {
      disposeStream(upstream.stream);
      return sendText(res, 410, 'Download URL expired. Request a new attachment download URL.');
    }

    const headers = {
      'Cache-Control': 'no-store',
      'Content-Type': upstream.headers['content-type'] || ticket.contentType || 'application/octet-stream',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeContentDispositionFilename(ticket.name)}`,
    };

    if (contentLength > 0) {
      headers['Content-Length'] = String(contentLength);
    }

    if (typeof res.set === 'function') {
      res.set(headers);
      res.status(200);
    } else {
      res.writeHead(200, headers);
    }

    upstream.stream.on('error', () => {
      if (typeof res.destroy === 'function') {
        res.destroy();
      }
    });

    upstream.stream.pipe(res);
  } catch (error) {
    return sendText(res, 502, 'Failed to download attachment from Microsoft Graph.');
  }
}

module.exports = {
  handleAttachmentDownloadRequest,
};

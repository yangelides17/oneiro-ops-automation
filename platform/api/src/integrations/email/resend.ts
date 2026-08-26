import { Resend } from 'resend';
import { config } from '../../config.js';
import type { EmailConnector, EmailParams } from './interface.js';

let client: Resend | null = null;

function getClient(): Resend {
  if (!client) {
    if (!config.resend.apiKey) throw new Error('RESEND_API_KEY not configured');
    client = new Resend(config.resend.apiKey);
  }
  return client;
}

export const resendEmail: EmailConnector = {
  async send(_orgId, params) {
    const to = Array.isArray(params.to) ? params.to : [params.to];

    await getClient().emails.send({
      from: config.resend.fromEmail,
      to,
      subject: params.subject,
      html: params.html,
      attachments: params.attachments?.map((a) => ({
        filename: a.filename,
        content: a.content,
      })),
    });
  },
};

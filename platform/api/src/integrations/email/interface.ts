export interface EmailParams {
  to: string | string[];
  subject: string;
  html: string;
  attachments?: { filename: string; content: Buffer }[];
}

export interface EmailConnector {
  send(orgId: string, params: EmailParams): Promise<void>;
}

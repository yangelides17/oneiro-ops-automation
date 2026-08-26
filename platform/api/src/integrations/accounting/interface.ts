/** Accounting integration interface — designed now, QuickBooks connector deferred. */

export interface InvoiceLineItem {
  description: string;
  quantity: number;
  unitPrice: number;
  amount: number;
  itemId?: string; // external system item ID
}

export interface InvoiceData {
  customerName: string;
  invoiceDate: string;   // YYYY-MM-DD
  dueDate: string;
  lines: InvoiceLineItem[];
  total: number;
  memo?: string;
}

export interface AccountingConnector {
  createInvoice(orgId: string, invoice: InvoiceData): Promise<{ externalId: string; docNumber: string }>;
  getInvoice(orgId: string, externalId: string): Promise<InvoiceData | null>;
  lookupCustomer(orgId: string, name: string): Promise<{ customerId: string } | null>;
  getInvoiceUrl(orgId: string, externalId: string): Promise<string>;
}

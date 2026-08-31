import { z } from 'zod';

// ─── Auth ────────────────────────────────────────────────────

export const signupSchema = z.object({
  email: z.string().email(),
  password: z.string().min(8, 'Password must be at least 8 characters'),
  name: z.string().min(1, 'Name is required'),
  orgName: z.string().min(1, 'Organization name is required'),
});

export const loginSchema = z.object({
  email: z.string().email(),
  password: z.string().min(1),
});

export const inviteSchema = z.object({
  email: z.string().email(),
  role: z.enum(['admin', 'foreman', 'crew']),
});

export const acceptInviteSchema = z.object({
  token: z.string().min(1),
  name: z.string().min(1),
  password: z.string().min(8),
});

// ─── Organization ────────────────────────────────────────────

export const updateOrgSchema = z.object({
  name: z.string().min(1).optional(),
  address: z.string().optional(),
  phone: z.string().optional(),
  email: z.string().email().optional(),
  taxId: z.string().optional(),
  timezone: z.string().optional(),
  opDayCutoffHour: z.number().int().min(0).max(23).optional(),
  signatoryName: z.string().optional(),
  signatoryTitle: z.string().optional(),
});

// ─── Work Order ──────────────────────────────────────────────

export const woStatusSchema = z.enum([
  'received', 'dispatched', 'in_progress', 'completed', 'returned',
]);

export const createWoSchema = z.object({
  woNumber: z.string().min(1),
  contractorId: z.string().uuid(),
  contractId: z.string().optional(),
  contractNum: z.string().optional(),
  regionCode: z.string().optional(),
  location: z.string().optional(),
  fromStreet: z.string().optional(),
  toStreet: z.string().optional(),
  dueDate: z.string().optional(),
  priority: z.string().optional(),
  workType: z.string().optional(),
  woReceivedDate: z.string().optional(),
  waterBlastRequired: z.string().optional(),
  notes: z.string().optional(),
  generalRemarks: z.string().optional(),
  school: z.string().optional(),
  prepBy: z.string().optional(),
  dateEntered: z.string().optional(),
  scanData: z.any().optional(),
});

// ─── Marking Item ────────────────────────────────────────────

export const createMarkingItemSchema = z.object({
  category: z.string().min(1),
  workType: z.string().optional(),
  woSection: z.enum(['top_table', 'intersection_grid', 'manual']).optional(),
  intersection: z.string().optional(),
  direction: z.string().optional(),
  description: z.string().optional(),
  quantity: z.number().optional(),
  unit: z.enum(['SF', 'LF', 'EA']).optional(),
  colorMaterial: z.string().optional(),
  crewChief: z.string().optional(),
  addedBy: z.enum(['scanner', 'manual']).optional(),
  notes: z.string().optional(),
});

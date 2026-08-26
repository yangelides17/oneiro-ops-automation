import { eq, and, asc } from 'drizzle-orm';
import type { Db } from '../client.js';
import {
  regions, contractors, markingCategories, pricingMultipliers,
  contractPricing, payClassifications, payRates, overtimeRules,
  billingRemaps, employees, contractLookup,
} from '../schema.js';

// ─── Regions ─────────────────────────────────────────────────

export async function listRegions(db: Db, orgId: string) {
  return db.select().from(regions)
    .where(eq(regions.orgId, orgId))
    .orderBy(asc(regions.sortOrder));
}

export async function createRegion(db: Db, orgId: string, data: { code: string; name: string; sortOrder?: number }) {
  const [row] = await db.insert(regions).values({ orgId, ...data }).returning();
  return row;
}

export async function updateRegion(db: Db, orgId: string, id: string, data: Partial<{ code: string; name: string; sortOrder: number }>) {
  const [row] = await db.update(regions).set(data)
    .where(and(eq(regions.id, id), eq(regions.orgId, orgId)))
    .returning();
  return row;
}

export async function deleteRegion(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(regions).where(and(eq(regions.id, id), eq(regions.orgId, orgId))).returning({ id: regions.id });
  return deleted.length > 0;
}

// ─── Contractors ─────────────────────────────────────────────

export async function listContractors(db: Db, orgId: string) {
  return db.select().from(contractors)
    .where(eq(contractors.orgId, orgId))
    .orderBy(asc(contractors.name));
}

export async function createContractor(db: Db, orgId: string, data: typeof contractors.$inferInsert) {
  const [row] = await db.insert(contractors).values({ ...data, orgId }).returning();
  return row;
}

export async function updateContractor(db: Db, orgId: string, id: string, data: Partial<typeof contractors.$inferInsert>) {
  const [row] = await db.update(contractors).set(data)
    .where(and(eq(contractors.id, id), eq(contractors.orgId, orgId)))
    .returning();
  return row;
}

export async function deleteContractor(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(contractors).where(and(eq(contractors.id, id), eq(contractors.orgId, orgId))).returning({ id: contractors.id });
  return deleted.length > 0;
}

// ─── Marking Categories ─────────────────────────────────────

export async function listCategories(db: Db, orgId: string) {
  return db.select().from(markingCategories)
    .where(eq(markingCategories.orgId, orgId))
    .orderBy(asc(markingCategories.sortOrder));
}

export async function createCategory(db: Db, orgId: string, data: typeof markingCategories.$inferInsert) {
  const [row] = await db.insert(markingCategories).values({ ...data, orgId }).returning();
  return row;
}

export async function updateCategory(db: Db, orgId: string, id: string, data: Partial<typeof markingCategories.$inferInsert>) {
  const [row] = await db.update(markingCategories).set(data)
    .where(and(eq(markingCategories.id, id), eq(markingCategories.orgId, orgId)))
    .returning();
  return row;
}

export async function deleteCategory(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(markingCategories).where(and(eq(markingCategories.id, id), eq(markingCategories.orgId, orgId))).returning({ id: markingCategories.id });
  return deleted.length > 0;
}

// ─── Pricing Multipliers ────────────────────────────────────

export async function listMultipliers(db: Db, orgId: string) {
  return db.select().from(pricingMultipliers)
    .where(eq(pricingMultipliers.orgId, orgId));
}

export async function upsertMultiplier(db: Db, orgId: string, data: { categoryName: string; multiplierType: string; value: string }) {
  const [row] = await db.insert(pricingMultipliers)
    .values({ orgId, ...data })
    .onConflictDoUpdate({
      target: [pricingMultipliers.orgId, pricingMultipliers.categoryName, pricingMultipliers.multiplierType],
      set: { value: data.value },
    })
    .returning();
  return row;
}

export async function deleteMultiplier(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(pricingMultipliers).where(and(eq(pricingMultipliers.id, id), eq(pricingMultipliers.orgId, orgId))).returning({ id: pricingMultipliers.id });
  return deleted.length > 0;
}

// ─── Contract Pricing ───────────────────────────────────────

export async function listContractPricing(db: Db, orgId: string) {
  return db.select().from(contractPricing)
    .where(eq(contractPricing.orgId, orgId))
    .orderBy(asc(contractPricing.contractNum));
}

export async function createContractPricing(db: Db, orgId: string, data: typeof contractPricing.$inferInsert) {
  const [row] = await db.insert(contractPricing).values({ ...data, orgId }).returning();
  return row;
}

export async function updateContractPricing(db: Db, orgId: string, id: string, data: Partial<typeof contractPricing.$inferInsert>) {
  const [row] = await db.update(contractPricing).set(data)
    .where(and(eq(contractPricing.id, id), eq(contractPricing.orgId, orgId)))
    .returning();
  return row;
}

export async function deleteContractPricing(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(contractPricing).where(and(eq(contractPricing.id, id), eq(contractPricing.orgId, orgId))).returning({ id: contractPricing.id });
  return deleted.length > 0;
}

// ─── Pay Classifications ────────────────────────────────────

export async function listClassifications(db: Db, orgId: string) {
  return db.select().from(payClassifications)
    .where(eq(payClassifications.orgId, orgId))
    .orderBy(asc(payClassifications.sortOrder));
}

export async function createClassification(db: Db, orgId: string, data: { code: string; name: string; sortOrder?: number }) {
  const [row] = await db.insert(payClassifications).values({ orgId, ...data }).returning();
  return row;
}

export async function deleteClassification(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(payClassifications).where(and(eq(payClassifications.id, id), eq(payClassifications.orgId, orgId))).returning({ id: payClassifications.id });
  return deleted.length > 0;
}

// ─── Pay Rates ──────────────────────────────────────────────

export async function listPayRates(db: Db, orgId: string) {
  return db.select().from(payRates)
    .where(eq(payRates.orgId, orgId))
    .orderBy(asc(payRates.classificationCode), asc(payRates.effectiveDate));
}

export async function createPayRate(db: Db, orgId: string, data: typeof payRates.$inferInsert) {
  const [row] = await db.insert(payRates).values({ ...data, orgId }).returning();
  return row;
}

export async function updatePayRate(db: Db, orgId: string, id: string, data: Partial<typeof payRates.$inferInsert>) {
  const [row] = await db.update(payRates).set(data)
    .where(and(eq(payRates.id, id), eq(payRates.orgId, orgId)))
    .returning();
  return row;
}

export async function deletePayRate(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(payRates).where(and(eq(payRates.id, id), eq(payRates.orgId, orgId))).returning({ id: payRates.id });
  return deleted.length > 0;
}

// ─── Overtime Rules ─────────────────────────────────────────

export async function getOvertimeRules(db: Db, orgId: string) {
  const [row] = await db.select().from(overtimeRules)
    .where(eq(overtimeRules.orgId, orgId))
    .limit(1);
  return row ?? null;
}

export async function updateOvertimeRules(db: Db, orgId: string, data: Partial<typeof overtimeRules.$inferInsert>) {
  const [row] = await db.update(overtimeRules).set(data)
    .where(eq(overtimeRules.orgId, orgId))
    .returning();
  return row;
}

// ─── Billing Remaps ─────────────────────────────────────────

export async function listBillingRemaps(db: Db, orgId: string) {
  return db.select().from(billingRemaps)
    .where(eq(billingRemaps.orgId, orgId));
}

export async function createBillingRemap(db: Db, orgId: string, data: typeof billingRemaps.$inferInsert) {
  const [row] = await db.insert(billingRemaps).values({ ...data, orgId }).returning();
  return row;
}

export async function deleteBillingRemap(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(billingRemaps).where(and(eq(billingRemaps.id, id), eq(billingRemaps.orgId, orgId))).returning({ id: billingRemaps.id });
  return deleted.length > 0;
}

// ─── Employees ──────────────────────────────────────────────

export async function listEmployees(db: Db, orgId: string) {
  return db.select().from(employees)
    .where(and(eq(employees.orgId, orgId), eq(employees.isActive, true)))
    .orderBy(asc(employees.name));
}

export async function createEmployee(db: Db, orgId: string, data: typeof employees.$inferInsert) {
  const [row] = await db.insert(employees).values({ ...data, orgId }).returning();
  return row;
}

export async function updateEmployee(db: Db, orgId: string, id: string, data: Partial<typeof employees.$inferInsert>) {
  const [row] = await db.update(employees).set(data)
    .where(and(eq(employees.id, id), eq(employees.orgId, orgId)))
    .returning();
  return row;
}

// ─── Contract Lookup ────────────────────────────────────────

export async function listContractLookup(db: Db, orgId: string) {
  return db.select().from(contractLookup)
    .where(eq(contractLookup.orgId, orgId));
}

export async function upsertContractLookup(db: Db, orgId: string, data: typeof contractLookup.$inferInsert) {
  const [row] = await db.insert(contractLookup)
    .values({ ...data, orgId })
    .onConflictDoUpdate({
      target: [contractLookup.orgId, contractLookup.contractNum, contractLookup.regionCode],
      set: { contractId: data.contractId, projectName: data.projectName, regionName: data.regionName },
    })
    .returning();
  return row;
}

export async function deleteContractLookup(db: Db, orgId: string, id: string) {
  const deleted = await db.delete(contractLookup).where(and(eq(contractLookup.id, id), eq(contractLookup.orgId, orgId))).returning({ id: contractLookup.id });
  return deleted.length > 0;
}

import { Router } from 'express';
import orgRouter from './org.js';
import regionsRouter from './regions.js';
import contractorsRouter from './contractors.js';
import categoriesRouter from './categories.js';
import pricingRouter from './pricing.js';
import payrollRouter from './payroll.js';
import employeesRouter from './employees.js';
import usersRouter from './users.js';
import billingRemapsRouter from './billingRemaps.js';
import contractLookupRouter from './contractLookup.js';

const router = Router();

router.use('/org', orgRouter);
router.use('/regions', regionsRouter);
router.use('/contractors', contractorsRouter);
router.use('/categories', categoriesRouter);
router.use('/pricing', pricingRouter);
router.use('/payroll', payrollRouter);
router.use('/employees', employeesRouter);
router.use('/users', usersRouter);
router.use('/billing-remaps', billingRemapsRouter);
router.use('/contract-lookup', contractLookupRouter);

export default router;

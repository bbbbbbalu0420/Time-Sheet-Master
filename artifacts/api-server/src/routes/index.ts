import { Router, type IRouter } from "express";
import healthRouter from "./health";
import payrollRouter from "./payroll";

const router: IRouter = Router();

router.use(healthRouter);
router.use(payrollRouter);

export default router;

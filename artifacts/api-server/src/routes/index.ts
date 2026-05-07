import { Router, type IRouter } from "express";
import healthRouter from "./health";
import fillXlsRouter from "./fillXls";

const router: IRouter = Router();

router.use(healthRouter);
router.use(fillXlsRouter);

export default router;

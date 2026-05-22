import { Router } from "express";
import multer from "multer";
import { spawn } from "child_process";
import { writeFileSync, unlinkSync, existsSync, readFileSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";
import { fileURLToPath } from "url";

const router = Router();
const routeDir = fileURLToPath(new URL(".", import.meta.url));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { 
    fileSize: 50 * 1024 * 1024,   // 50MB max file size (purani setting)
    fieldSize: 25 * 1024 * 1024,  // 25MB max text field size (Nayi line: Taaki dataRows crash na karein)
  },
});

router.post(
  "/fill-xls",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "oldCatalog", maxCount: 1 },
  ]),
  async (req, res) => {
    const id = randomUUID();
    const tmp = tmpdir();
    const templatePath  = join(tmp, `tmpl_${id}.xls`);
    const oldPath       = join(tmp, `old_${id}.xls`);
    const mappingsPath  = join(tmp, `map_${id}.json`);
    const configPath    = join(tmp, `cfg_${id}.json`);
    const outputPath    = join(tmp, `out_${id}.xls`);

    try {
      const files = req.files as Record<string, Express.Multer.File[]>;
      const templateFile  = files?.["template"]?.[0];
      const oldFile       = files?.["oldCatalog"]?.[0];
      const mappingsJson  = req.body?.mappings;

      if (!templateFile || !oldFile || !mappingsJson) {
        res.status(400).json({ error: "template, oldCatalog, and mappings are required" });
        return;
      }

      writeFileSync(templatePath, templateFile.buffer);
      writeFileSync(oldPath, oldFile.buffer);
      writeFileSync(mappingsPath, mappingsJson);

      // Config JSON — carries exact header row indices detected by the frontend
      const config: Record<string, unknown> = {
        templatePath,
        oldPath,
        outputPath,
        mappingsPath,
        templateSheetIdx:  Number(req.body?.templateSheetIdx  ?? 0),
        templateHeaderRow: Number(req.body?.templateHeaderRow ?? -1),
        oldSheetIdx:       Number(req.body?.oldSheetIdx       ?? 0),
        oldHeaderRow:      Number(req.body?.oldHeaderRow      ?? -1),
      };
      // If the frontend sent pre-processed rows (prefix/suffix already applied),
      // embed them in the config so Python uses them directly instead of
      // re-reading from the raw old catalog file.
      if (req.body?.dataRows) {
        try {
          config.dataRows = JSON.parse(req.body.dataRows);
        } catch {
          // ignore parse errors — Python will fall back to reading old catalog
        }
      }
      writeFileSync(configPath, JSON.stringify(config));

      const pythonPath = process.env.PYTHON_PATH || "python3";
      const scriptPathCandidates = [
        join(process.cwd(), "../python/fill_xls.py"),
        join(process.cwd(), "python/fill_xls.py"),
        join(routeDir, "../../../python/fill_xls.py"),
        join(routeDir, "../../python/fill_xls.py"),
      ];
      const scriptPath =
        scriptPathCandidates.find((candidate) => existsSync(candidate)) ??
        scriptPathCandidates[0];

      // Include project-local .pythonlibs if present (local dev compatibility),
      // while still working when Railway installs packages globally.
      const pythonLibCandidates = [
        join(process.cwd(), "../.pythonlibs/lib/python3.11/site-packages"),
        join(process.cwd(), ".pythonlibs/lib/python3.11/site-packages"),
        join(routeDir, "../../../.pythonlibs/lib/python3.11/site-packages"),
        join(routeDir, "../../.pythonlibs/lib/python3.11/site-packages"),
      ];
      const pythonLibs = pythonLibCandidates.find((candidate) => existsSync(candidate));
      const existingPythonPath = process.env.PYTHONPATH || "";
      const spawnEnv = {
        ...process.env,
        PYTHONPATH: pythonLibs
          ? (existingPythonPath ? `${pythonLibs}:${existingPythonPath}` : pythonLibs)
          : existingPythonPath,
      };

      const { stdout, stderr } = await new Promise<{ stdout: string; stderr: string }>(
        (resolve, reject) => {
          const proc = spawn(pythonPath, [scriptPath, configPath], { env: spawnEnv });
          let stdout = "";
          let stderr = "";
          proc.stdout.on("data", (d: Buffer) => { stdout += d.toString(); });
          proc.stderr.on("data", (d: Buffer) => { stderr += d.toString(); });
          proc.on("close", (code: number) => {
            if (code !== 0) reject(new Error(`Python error (${code}): ${stderr}`));
            else resolve({ stdout, stderr });
          });
        }
      );

      req.log.info({ pythonDebug: stderr.trim() }, "Python fill_xls debug");

      const parsed = JSON.parse(stdout.trim());
      if (!parsed.success) throw new Error("Python script failed");

      const outputBuffer = readFileSync(outputPath);
      const originalName = templateFile.originalname || "flipkart_catalog_filled.xls";

      res.setHeader("Content-Type", "application/vnd.ms-excel");
      res.setHeader("Content-Disposition", `attachment; filename="${originalName}"`);
      res.setHeader("X-Rows-Written",  String(parsed.rowsWritten));
      res.setHeader("X-Cells-Written", String(parsed.cellsWritten));
      res.send(outputBuffer);
    } catch (err) {
      req.log.error(err);
      res.status(500).json({ error: String(err) });
    } finally {
      for (const p of [templatePath, oldPath, mappingsPath, configPath, outputPath]) {
        if (existsSync(p)) { try { unlinkSync(p); } catch {} }
      }
    }
  }
);

export default router;

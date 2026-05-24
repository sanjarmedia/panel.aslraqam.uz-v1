import { Router } from "express";
import { getCodesInfo } from "../services/codesInfoService.js";

const router = Router();

router.post("/info", async (req, res) => {
  try {
    const { codes, pg } = req.body;
    const data = await getCodesInfo(codes, pg);
    res.json(data);
  } catch (error) {
    res.status(500).json({ message: "Server error!" });
  }
});

export default router;

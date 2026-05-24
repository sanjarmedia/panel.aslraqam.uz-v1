import { Router } from "express";
import { getCodesHistory } from "../services/codesHistoryService.js";

const router = Router();

router.get("/history", async (req, res) => {
  try {
    const { id } = req.query;
    const data = await getCodesHistory(id);
    res.json(data);
  } catch (error) {
    res.status(500).json({ message: "Server error!" });
  }
});

export default router;

import express from "express";
import cors from "cors";
import dotenv from "dotenv";

import codesHistory from "./routes/codesHistory.js";
import codesInfo from "./routes/codesInfo.js";

const app = express();
app.use(express.json());
app.use(cors());
dotenv.config();

app.use("/api", codesHistory);
app.use("/api", codesInfo);

app.listen(8080, () => {
  console.log(`Server is running on port 8080`);
});

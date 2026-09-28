import "dotenv/config";
import { deviceCodeLogin } from "./auth.js";

deviceCodeLogin().catch((err) => {
  console.error(err);
  process.exit(1);
});

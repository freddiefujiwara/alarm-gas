import fs from "fs";
import os from "os";
import path from "path";
import { describe, expect, it } from "vitest";

import { build } from "../build";

describe("build", () => {
  it("writes dist/Code.gs without export statements and copies the manifest", () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "alarm-gas-build-"));
    fs.mkdirSync(path.join(root, "src"));
    fs.writeFileSync(
      path.join(root, "src", "Code.js"),
      "export const a = 1;\nconst b = 2;\n"
    );
    fs.writeFileSync(path.join(root, "appsscript.json"), "{}");
    build({ rootDir: root });
    const out = fs.readFileSync(path.join(root, "dist", "Code.gs"), "utf8");
    expect(out).not.toContain("export");
    expect(out).toContain("const a = 1;");
    expect(fs.existsSync(path.join(root, "dist", "appsscript.json"))).toBe(true);
  });
});

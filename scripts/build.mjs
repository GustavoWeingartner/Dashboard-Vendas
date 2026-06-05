import fs from "node:fs/promises";
import vm from "node:vm";

const babelCode = await fs.readFile("public/vendor/babel.min.js", "utf8");
const context = {};
vm.createContext(context);
vm.runInContext(babelCode, context);

const source = await fs.readFile("src/app.tsx", "utf8");
const compiled = context.Babel.transform(source, {
  filename: "app.tsx",
  presets: ["typescript", "react"],
  comments: true,
}).code;

await fs.writeFile(
  "public/app.compiled.js",
  `${compiled}\n//# sourceURL=src/app.tsx\n`,
  "utf8",
);

await fs.copyFile("index.html", "public/index.html");

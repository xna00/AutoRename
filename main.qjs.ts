// @ts-nocheck

import * as std from "std";
import * as os from "os";
import { DOMParser, serializeToWellFormedString } from "slimdom";
import UZIP from "uzip";
import parseArgs from "minimist";

const exec = (cmd: string) => {
  const err = {};
  console.log("run: ", cmd);
  const f = std.popen(cmd, "r", err);
  if (!f) {
    throw new Error(`Can not run: ${cmd}, errno is ${err}`);
  }
  console.log("done: ", cmd);
  console.log("f: ", f);
  const ret = f.readAsString();
  console.log("ret: ", ret);
  return ret;
};

const readFileSync = (path: string) => {
  const f = std.open(path, "rb+");
  const ret: number[] = [];
  while (true) {
    const b = f.getByte();
    if (b === -1) {
      break;
    }
    ret.push(b);
  }
  f.close();
  return new Uint8Array(ret);
};
const renameSync = (oldPath: string, newPath: string) => {
  return os.rename(oldPath, newPath);
};

const dirname = (path: string) => path.split("\\").slice(0, -1).join("\\");

const assert: <T>(
  c: T,
  m: string
) => asserts c is Exclude<NonNullable<T>, false> = (c, msg) => {
  if (!c) {
    throw new Error(msg);
  }
};

const bufferToString = (buf: Uint8Array) => {
  const f = std.tmpfile();
  f.write(buf.buffer, 0, buf.length);
  f.flush();
  f.seek(0, 0);
  const ret = f.readAsString();
  f.close();
  return ret;
};

const [__filename] = os.realpath(globalThis.scriptArgs[0]);
console.log(__filename);
const args = parseArgs(globalThis.scriptArgs.slice(1));
const info = `
1: 设置注册表
2: 取消设置注册表
3: 跳过小字
4: 设置最小字号
q: 退出

`.trimStart();
const unsetRegistry = () => {
  let oldEntry = "";
  try {
    const t = exec("reg query HKCU\\Software\\AutoRename /v Entry");
    const e = t.split("REG_SZ")[1].trim();
    oldEntry = e;
  } catch {}
  if (!oldEntry) {
    return;
  }
  try {
    exec(`reg delete HKCR\\${oldEntry}\\shell\\AutoRename /f`);
    exec("reg delete HKCU\\Software\\AutoRename /f");
  } catch {}
};
if (!args._[0]) {
  while (true) {
    console.log(info);
    const c = std.getche();
    const s = String.fromCharCode(c);
    if (s === "1") {
      unsetRegistry();
      console.log("设置注册表");
      const assoc = exec("assoc .docx");
      console.log("assoc: ", assoc);
      const entry = assoc.split("=")[1].trim();
      console.log("entry: ", entry);
      exec(`reg add HKCR\\${entry}\\shell\\AutoRename /ve /d "自动重命名" /f`);

      exec(
        `reg add HKCR\\${entry}\\shell\\AutoRename\\command /ve /d """${__filename}"" ""%1""" /f`
      );
      exec(`reg add HKCU\\Software\\AutoRename /v Entry /d ${entry} /f`);
    } else if (s === "2") {
      console.log("unset");
      unsetRegistry();
    } else if (s === "3") {
      console.log("skip");
    } else if (s === "4") {
      console.log("min");
    } else if (s === "q") {
      std.exit(0);
    } else {
      console.log("输入错误");
    }
  }
}

const [filePath, err] = os.realpath(args._[0]);

if (!filePath) {
  console.log("请提供 .docx 文件路径作为参数");
  std.exit(1);
}
console.log("filePath", filePath);

const file = readFileSync(filePath);
const files = UZIP.parse(file);

console.log("files", Object.keys(files));

const documentXml = bufferToString(files["word/document.xml"]);
const stylesXml = bufferToString(files["word/styles.xml"]);

assert(documentXml, "无法找到 word/document.xml");
assert(stylesXml, "无法找到 word/styles.xml");

const styles = new DOMParser().parseFromString(stylesXml, "text/xml");
const styleList = styles.getElementsByTagName("w:style");

const getStyleById = (id: string) =>
  styleList.find((s) => s.getAttribute("w:styleId") === id);

const paragraphs = new DOMParser()
  .parseFromString(documentXml, "text/xml")
  .getElementsByTagName("w:body")[0]
  .children.filter((n) => n.nodeName === "w:p");

// console.log("正在解析文件:", paragraphs.length, "个段落");

let title = "",
  fontSize = 0;

for (const p of paragraphs) {
  const wts = p.getElementsByTagName("w:t");
  const text = wts.map((wt) => wt.textContent?.trim() || "").join("");
  // console.log(title, text);
  if (!title && !text) {
    continue;
  } else {
    if (!text) break;
    const wr = p.getElementsByTagName("w:r").at(0);
    assert(wr, "段落内没有文本");
    let size = wr.getElementsByTagName("w:sz").at(0)?.getAttribute("w:val");
    if (!size) {
      const ppr = p.getElementsByTagName("w:pPr").at(0);
      if (ppr) {
        size = ppr.getElementsByTagName("w:sz").at(0)?.getAttribute("w:val");
      }
    }
    if (!size) {
      const ppr = p.getElementsByTagName("w:pPr").at(0);
      if (ppr) {
        const pStyleId = ppr
          .getElementsByTagName("w:pStyle")
          .at(0)
          ?.getAttribute("w:val");
        if (pStyleId) {
          const style = getStyleById(pStyleId);
          if (style) {
            size = style
              .getElementsByTagName("w:rPr")
              .at(0)
              ?.getElementsByTagName("w:sz")
              .at(0)
              ?.getAttribute("w:val");
          }
        }
      }
    }
    if (!size || parseInt(size) < 28) break;
    if (!fontSize) {
      fontSize = parseInt(size);
    }
    if (fontSize !== parseInt(size)) break;
    else title += p.textContent?.trim() || "";
  }
}

// console.log("正在解析文件:", styles, getStyleById("2"));
console.log("提取到的标题:", title, "字号:", fontSize);
if (title) {
  const dir = dirname(filePath);
  renameSync(filePath, `${dir}\\${title}.docx`);
}

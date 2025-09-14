#! /usr/bin/env node

import { DOMParser, serializeToWellFormedString } from "slimdom";
// import jszip from "jszip";
import { readFileSync, renameSync } from "fs";
import assert from "assert";
import { dirname } from "path";
import UZIP from "uzip";

const filePath = process.argv[2];
if (!filePath) {
  console.error("请提供 .docx 文件路径作为参数");
  process.exit(1);
}

const file = readFileSync(filePath);
const files = UZIP.parse(file.buffer);
const documentXml = Buffer.from(files["word/document.xml"]).toString();
const stylesXml = Buffer.from(files["word/styles.xml"]).toString();

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
    const wr = p.getElementsByTagName("w:r")[0];
    assert(wr, "段落内没有文本");
    let size = wr.getElementsByTagName("w:sz")[0]?.getAttribute("w:val");
    if (!size) {
      const ppr = p.getElementsByTagName("w:pPr")[0];
      if (ppr) {
        size = ppr.getElementsByTagName("w:sz")[0]?.getAttribute("w:val");
      }
    }
    if (!size) {
      const ppr = p.getElementsByTagName("w:pPr")[0];
      if (ppr) {
        const pStyleId = ppr
          .getElementsByTagName("w:pStyle")[0]
          ?.getAttribute("w:val");
        if (pStyleId) {
          const style = getStyleById(pStyleId);
          if (style) {
            size = style
              .getElementsByTagName("w:rPr")[0]
              .getElementsByTagName("w:sz")[0]
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
console.log("提取到的标题:", title, fontSize);
if (title) {
  const dir = dirname(filePath);
  renameSync(filePath, `${dir}/${title}.docx`);
}

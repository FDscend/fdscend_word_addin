function adjStyle() {
  // 获取用户的首选语言
  const userLanguage = navigator.language.toLowerCase();

  // 判断是否是简体中文 (zh-cn)
  if (userLanguage === "zh-cn") {
    // 如果是简体中文，添加对应的 CSS 规则
    const style = document.createElement("style");
    style.innerHTML = `
        h1::before {
            content: counter(chapterH1, cjk-ideographic) '\\3001';
        }
    `;
    document.head.appendChild(style);
  } else {
    // 如果不是简体中文，添加默认的 CSS 规则
    const style = document.createElement("style");
    style.innerHTML = `
        h1::before {
            content: counter(chapterH1) " ";
        }
    `;
    document.head.appendChild(style);
  }
}

function convertMarkdownWithEmoji(html) {
  // 1. 先用marked转换Markdown

  // 2. 使用正则表达式保护代码块
  html = html.replace(
    /<pre><code>[\s\S]*?<\/code><\/pre>|<code>.*?<\/code>/g,
    (match) => match.replace(/:/g, "&#58;")
  );

  // 3. 转换emoji
  html = joypixels.toImage(html);

  // 4. 恢复代码块中的冒号
  html = html.replace(/&#58;/g, ":");

  return html;
}

function parseFrontMatter(content) {
  const fmRegex = /^---\n([\s\S]+?)\n---\n/;
  const match = content.match(fmRegex);

  if (match) {
    try {
      const yaml = match[1];
      const data = jsyaml.load(yaml);
      const mdContent = content.slice(match[0].length);
      return { data, content: mdContent };
    } catch (e) {
      console.error("YAML 解析错误:", e);
      return { data: {}, content };
    }
  }
  return { data: {}, content };
}

function parseFrontMatter2(content) {
  const fmRegex = /^---\n([\s\S]+?)\n---\n/;
  const match = content.match(fmRegex);

  if (match) {
    try {
      const yaml = match[1];
      const mdContent = content.slice(match[0].length);
      return { data: yaml, content: mdContent };
    } catch (e) {
      console.error("YAML 解析错误:", e);
      return { data: {}, content };
    }
  }
  return { data: {}, content };
}

function renderMetadataValue(value) {
  if (Array.isArray(value)) {
    if (value.length === 0) return '<span class="metadata-value">-</span>';

    // 特殊处理 tags 数组
    if (value.every((item) => typeof item === "string")) {
      return `
        <div class="metadata-value">
          <div class="tags-container">
              ${value
                .map((tag) => `<span class="front-tag">${tag}</span>`)
                .join("")}
          </div>
        </div>
      `;
    }

    return `
      <div class="metadata-value">
        ${value
          .map((item) => `<div>${renderMetadataValue(item)}</div>`)
          .join("")}
        </div>
      `;
  }

  if (typeof value === "object" && value !== null) {
    return `
      <div class="metadata-value">
        ${Object.entries(value)
          .map(
            ([k, v]) => `
            <div class="metadata-row">
                <div class="metadata-key">${k}:</div>
                ${renderMetadataValue(v)}
            </div>
        `
          )
          .join("")}
      </div>
    `;
  }

  // 特殊处理日期
  if (typeof value === "string" && value.match(/^\d{4}-\d{2}-\d{2}/)) {
    return `<span class="metadata-value date-value">${value}</span>`;
  }

  return `<span class="metadata-value">${value}</span>`;
}

function renderFrontMatter(data) {
  if (!data || Object.keys(data).length === 0) return "";

  return `
      <div class="front-matter-container">
          <div class="front-matter-title">文档属性</div>
          ${Object.entries(data)
            .map(
              ([key, value]) => `
              <div class="metadata-row">
                  <div class="metadata-key">${key}:</div>
                  ${renderMetadataValue(value)}
              </div>
          `
            )
            .join("")}
      </div>
  `;
}

function renderFrontMatter2(data) {
  if (!data) return "";

  return `<pre><code class="hljs language-yaml">${data}</code></pre>`;
}

const urlEncodedHeadingExtension = {
  name: "urlEncodedHeading",
  level: "block",
  start(src) {
    // 提示：可能的起点（提高性能）
    return src.match(/^#{1,6}\s/m)?.index;
  },
  tokenizer(src) {
    // 只处理 ATX 标题（# 到 ######），不处理 setext（===/---）标题
    const rule = /^(#{1,6})[ \t]+([^\n]+?)[ \t]*(?:\n+|$)/;
    const match = rule.exec(src);
    if (!match) return;

    const depth = match[1].length;
    const raw = match[0];
    const text = match[2];

    // 让标题里的行内 Markdown（如 **粗体**）继续生效
    const tokens = this.lexer.inlineTokens(text);

    return {
      type: "urlEncodedHeading",
      raw,
      depth,
      text,
      tokens,
    };
  },
  renderer(token) {
    const innerHTML = this.parser.parseInline(token.tokens);
    // 用“纯文本”来生成 id（避免把 <em> 等标签计入）
    const id = encodeURIComponent(token.text.trim().replace(/\s+/g, "-").toLowerCase());
    return `<h${token.depth} id="${id}">${innerHTML}</h${token.depth}>`;
  },
};

const mermaidExtension = {
  name: "mermaid",
  level: "block",
  start(src) {
    return src.match(/^```mermaid/)?.index;
  },
  tokenizer(src, tokens) {
    const rule = /^```mermaid([\s\S]*?)```/;
    const match = rule.exec(src);
    if (match) {
      return {
        type: "mermaid",
        raw: match[0],
        text: match[1].trim(),
      };
    }
  },
  renderer(token) {
    return `<pre class="mermaid">${token.text}</pre>`;
  },
};

const admonitionExtension = {
  name: "admonition",
  level: "block",
  start(src) {
    return src.match(/^\s{0,3}>\s+\[!/m)?.index;
  },
  tokenizer(src, tokens) {
    const rule = /^((?:\s{0,3}>\s?.*\n?)*)/;
    const match = rule.exec(src);
    if (!match) return;

    const block = match[1];
    const lines = block.split(/\n/);

    // 第一行
    const firstLine = lines[0].replace(/^\s{0,3}>\s?/, "");
    const headMatch = /^\[!([A-Za-z0-9]+(?:\|[^\]]*)?)\]([+-]?)(.*)/.exec(
      firstLine
    );
    if (!headMatch) return;

    const rawKeyword = headMatch[1];
    const keyword = rawKeyword.split("|")[0].toUpperCase();
    const collapseFlag = headMatch[2];
    const customTitle = headMatch[3].trim();

    const bodyLines = [];
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i];
      if (/^\s*$/.test(line)) break;
      const content = line.replace(/^\s{0,3}>\s?/, "");
      bodyLines.push(content);
    }

    const body = bodyLines.join("\n");

    return {
      type: "admonition",
      raw: lines.slice(0, bodyLines.length + 1).join("\n") + "\n",
      keyword,
      title:
        customTitle ||
        thmMeta[keyword]?.defaultTitle ||
        alertMeta[keyword]?.defaultTitle ||
        keyword,
      collapseState: collapseFlag,
      tokens: this.lexer.blockTokens(body),
    };
  },
  renderer(token) {
    if (thmMeta[token.keyword]) {
      // 定理类：白色背景 + 黑色边框
      if (token.collapseState) {
        const openAttr = token.collapseState === "+" ? " open" : "";

        return `
          <details class="thm-block"${openAttr}>
            <summary style="display:flex; align-items:center; font-weight:bold; cursor:pointer; list-style:none;">
              <span class="title-text">${
                thmMeta[token.keyword]["defaultTitle"]
              } (</span><span style="font-weight: normal;">${
          token.title
        }</span><span class="title-text">)</span>
              <span class="arrow"></span>
            </summary>
            <div class="alert-body" style="margin-top:0.6em;">
              ${this.parser.parse(token.tokens)}
            </div>
          </details>`;
      } else {
        return `
          <div class="thm-block">
            <div style="font-weight:bold; margin-bottom:0.4em;">
              <span class="title-text">${
                thmMeta[token.keyword]["defaultTitle"]
              } (</span><span style="font-weight: normal;">${
          token.title
        }</span><span class="title-text">)</span>
            </div>
            <div class="alert-body">
              ${this.parser.parse(token.tokens)}
            </div>
          </div>`;
      }
    } else {
      // 默认走原有 alertMeta 渲染（有颜色，有图标）
      const meta = alertMeta[token.keyword] || alertMeta["NOTE"];
      const icon = meta.icon;
      const color = meta.color;

      if (token.collapseState) {
        const openAttr = token.collapseState === "+" ? " open" : "";
        return `
          <details class="alert-block"${openAttr} style="background:${color}20;">
            <summary style="display:flex; align-items:center; font-weight:bold; color:${color}; cursor:pointer; list-style:none;">
              <span style="margin-right:0.4em;">${icon}</span>
              <span class="title-text">${token.title}</span>
              <span class="arrow"></span>
            </summary>
            <div class="alert-body" style="margin-top:0.6em;">
              ${this.parser.parse(token.tokens)}
            </div>
          </details>`;
      } else {
        return `
          <div class="alert-block" style="background:${color}20;">
            <div style="display:flex; align-items:center; font-weight:bold; color:${color}; margin-bottom:0.4em;">
              <span style="margin-right:0.4em;">${icon}</span> ${token.title}
            </div>
            <div class="alert-body">
              ${this.parser.parse(token.tokens)}
            </div>
          </div>`;
      }
    }
  },
};

const thmMeta = {
  THM: { defaultTitle: "Theorem" },
  DEF: { defaultTitle: "Definition" },
  LEM: { defaultTitle: "Lemma" },
};

const alertMeta = {
  NOTE: {
    defaultTitle: "Note",
    color: "#1775d9",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622l4.353-1.32a2 2 0 0 0 .83-.497z"></path><path d="m15 5 4 4"></path></svg>',
  },
  ABSTRACT: {
    defaultTitle: "Abstract",
    color: "#16a6ab",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path><path d="M12 11h4"></path><path d="M12 16h4"></path><path d="M8 11h.01"></path><path d="M8 16h.01"></path></svg>',
  },
  SUMMARY: {
    defaultTitle: "Abstract",
    color: "#16a6ab",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path><path d="M12 11h4"></path><path d="M12 16h4"></path><path d="M8 11h.01"></path><path d="M8 16h.01"></path></svg>',
  },
  INFO: {
    defaultTitle: "Info",
    color: "#1775d9",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M12 16v-4"></path><path d="M12 8h.01"></path></svg>',
  },
  TODO: {
    defaultTitle: "Todo",
    color: "#1775d9",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="m9 12 2 2 4-4"></path></svg>',
  },
  TIP: {
    defaultTitle: "Tip",
    color: "#16a6ab",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"></path></svg>',
  },
  HINT: {
    defaultTitle: "Hint",
    color: "#16a6ab",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"></path></svg>',
  },
  IMPORTANT: {
    defaultTitle: "Important",
    color: "#16a6ab",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"></path></svg>',
  },
  SUCCESS: {
    defaultTitle: "Success",
    color: "#1da51d",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"></path></svg>',
  },
  CHECK: {
    defaultTitle: "Check",
    color: "#1da51d",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"></path></svg>',
  },
  DONE: {
    defaultTitle: "Done",
    color: "#1da51d",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"></path></svg>',
  },
  QUESTION: {
    defaultTitle: "Question",
    color: "#de7417",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"></path><path d="M12 17h.01"></path></svg>',
  },
  HELP: {
    defaultTitle: "Help",
    color: "#de7417",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"></path><path d="M12 17h.01"></path></svg>',
  },
  FAQ: {
    defaultTitle: "Faq",
    color: "#de7417",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"></path><path d="M12 17h.01"></path></svg>',
  },
  WARNING: {
    defaultTitle: "Warning",
    color: "#de7417",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"></path><path d="M12 9v4"></path><path d="M12 17h.01"></path></svg>',
  },
  CAUTION: {
    defaultTitle: "Caution",
    color: "#de7417",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"></path><path d="M12 9v4"></path><path d="M12 17h.01"></path></svg>',
  },
  ATTENTION: {
    defaultTitle: "Attention",
    color: "#de7417",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"></path><path d="M12 9v4"></path><path d="M12 17h.01"></path></svg>',
  },
  FAILURE: {
    defaultTitle: "Failure",
    color: "#dd2c38",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18"></path><path d="m6 6 12 12"></path></svg>',
  },
  FAIL: {
    defaultTitle: "Fail",
    color: "#dd2c38",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18"></path><path d="m6 6 12 12"></path></svg>',
  },
  MISSING: {
    defaultTitle: "Missing",
    color: "#dd2c38",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18"></path><path d="m6 6 12 12"></path></svg>',
  },
  DANGER: {
    defaultTitle: "Danger",
    color: "#dd2c38",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z"></path></svg>',
  },
  ERROR: {
    defaultTitle: "Error",
    color: "#dd2c38",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z"></path></svg>',
  },
  BUG: {
    defaultTitle: "Bug",
    color: "#dd2c38",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m8 2 1.88 1.88"></path><path d="M14.12 3.88 16 2"></path><path d="M9 7.13v-1a3.003 3.003 0 1 1 6 0v1"></path><path d="M12 20c-3.3 0-6-2.7-6-6v-3a4 4 0 0 1 4-4h4a4 4 0 0 1 4 4v3c0 3.3-2.7 6-6 6"></path><path d="M12 20v-9"></path><path d="M6.53 9C4.6 8.8 3 7.1 3 5"></path><path d="M6 13H2"></path><path d="M3 21c0-2.1 1.7-3.9 3.8-4"></path><path d="M20.97 5c0 2.1-1.6 3.8-3.5 4"></path><path d="M22 13h-4"></path><path d="M17.2 17c2.1.1 3.8 1.9 3.8 4"></path></svg>',
  },
  EXAMPLE: {
    defaultTitle: "Example",
    color: "#8f47e1",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="8" y1="6" x2="21" y2="6"></line><line x1="8" y1="12" x2="21" y2="12"></line><line x1="8" y1="18" x2="21" y2="18"></line><line x1="3" y1="6" x2="3.01" y2="6"></line><line x1="3" y1="12" x2="3.01" y2="12"></line><line x1="3" y1="18" x2="3.01" y2="18"></line></svg>',
  },
  QUOTE: {
    defaultTitle: "Quote",
    color: "#9e9e9e",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"></path><path d="M5 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"></path></svg>',
  },
  CITE: {
    defaultTitle: "Cite",
    color: "#9e9e9e",
    icon: '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"></path><path d="M5 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"></path></svg>',
  },
};

const mathExtension = {
  name: "math",
  level: "inline", // 行内扩展
  start(src) {
    return src.match(/\$/)?.index;
  },
  tokenizer(src, tokens) {
    const match = src.match(/^\$\$([^$]*)\$\$/) || src.match(/^\$([^$]*)\$/);
    if (match) {
      return {
        type: "math",
        raw: match[0],
        text: match[1].trim(),
        display: match[0].startsWith("$$"), // 是否是行间公式
      };
    }
  },
  renderer(token) {
    try {
      return katex.renderToString(token.text, {
        throwOnError: false,
        displayMode: token.display,
      });
    } catch (error) {
      console.error("KaTeX 渲染错误:", error);
      return `<span class="katex-error" style="color:red">${token.raw}</span>`;
    }
  },
};

// 代码高亮扩展
const highlightCodeExtension = {
  name: "highlight",
  level: "block",
  start(src) {
    return src.match(/^```/)?.index;
  },
  tokenizer(src, tokens) {
    const rule = /^```(\w*)\n([\s\S]*?)\n```/;
    const match = rule.exec(src);
    if (match) {
      const lang = match[1].toLowerCase();
      // 识别mermaid代码块
      if (lang === "mermaid") {
        return {
          type: "mermaid",
          raw: match[0],
          text: match[2],
        };
      }
      return {
        type: "code",
        raw: match[0],
        lang: match[1],
        text: match[2],
      };
    }
  },
  renderer(token) {
    try {
      if (token.type === "mermaid") {
        return `<pre class="mermaid">${token.text}</pre>`;
      }
    } catch (e) {
      console.log("mermaid ext is not active", e);
    }

    // 自动检测语言
    let detectedLang = token.lang;
    if (
      !detectedLang ||
      !renderConfig.highlightLanguages.includes(detectedLang)
    ) {
      const autoDetect = hljs.highlightAuto(
        token.text,
        renderConfig.highlightLanguages
      );
      detectedLang = autoDetect.language || "plaintext";
    }

    // 应用用户选择的主题
    const langClass = detectedLang ? `language-${detectedLang}` : "";
    let highlighted;
    try {
      if (detectedLang && hljs.getLanguage(detectedLang)) {
        highlighted = hljs.highlight(token.text, {
          language: detectedLang,
        }).value;
      } else {
        highlighted = hljs.highlightAuto(token.text).value;
      }
    } catch (e) {
      console.warn("代码高亮错误:", e);
      highlighted = hljs.highlightAuto(token.text).value;
    }

    return `<pre><code class="hljs ${langClass}">${highlighted}</code></pre>`;
  },
};

const highlightExtension = {
  name: "highlight",
  level: "inline", // 行内级别扩展
  start(src) {
    return src.indexOf("=="); // 检测起始标记
  },
  tokenizer(src, tokens) {
    const rule = /^==([^=]+)==/; // 匹配 ==高亮内容==
    const match = rule.exec(src);
    if (match) {
      return {
        type: "highlight",
        raw: match[0],
        text: match[1].trim(),
      };
    }
  },
  renderer(token) {
    return `<mark class="text-highlight">${token.text}</mark>`;
  },
};

function renderMarkdown(content) {
  configureMarked();

  if (renderConfig.renderFront) {
    const { data, content: mdContent } = parseFrontMatter(content);
    const fmHtml = renderFrontMatter(data);

    return fmHtml + convertMarkdownWithEmoji(marked.parse(mdContent));
  } else {
    const { data, content: mdContent } = parseFrontMatter2(content);
    const fmHtml = renderFrontMatter2(data);

    return fmHtml + convertMarkdownWithEmoji(marked.parse(mdContent));
  }
}

function renderAll() {
  document.getElementById("md-body").innerHTML = renderMarkdown(
    document.getElementById("text").value
  );

  if (renderConfig.renderMermaid) {
    mermaid.run();
    document.querySelectorAll("pre code:not(.mermaid)").forEach((block) => {
      hljs.highlightElement(block);
    });
  }
}

const renderConfig = {
  renderAdmonition: true,
  renderMermaid: true,
  renderFootnote: true,
  renderMath: true,
  renderHighlight: true,
  renderFront: true,
  highlightTheme: "github", // 默认主题
};

// 初始化 marked 配置
function configureMarked() {
  // 重置 marked 配置
  marked.setOptions(marked.getDefaults());

  marked.setOptions({
    mangle: false,
    headerIds: false,
  });

  // 始终启用的扩展
  const alwaysEnabledExtensions = [urlEncodedHeadingExtension];

  // 条件启用的扩展
  const conditionalExtensions = [];
  if (renderConfig.renderAdmonition)
    conditionalExtensions.push(admonitionExtension);
  if (renderConfig.renderMermaid) {
    conditionalExtensions.push(mermaidExtension);
    conditionalExtensions.push(highlightCodeExtension);
  }
  if (renderConfig.renderMath) conditionalExtensions.push(mathExtension);
  if (renderConfig.renderFootnote) marked.use(markedFootnote());
  if (renderConfig.renderHighlight)
    conditionalExtensions.push(highlightExtension);

  // 应用扩展
  marked.use({
    extensions: [...alwaysEnabledExtensions, ...conditionalExtensions],
  });
}

// 加载保存的设置
function loadSettings() {
  const saved = localStorage.getItem("markdownRenderSettings");
  if (saved) {
    Object.assign(renderConfig, JSON.parse(saved));
    document.getElementById("renderAdmonition").checked =
      renderConfig.renderAdmonition;
    document.getElementById("renderMermaid").checked =
      renderConfig.renderMermaid;
    document.getElementById("renderMath").checked = renderConfig.renderMath;
    document.getElementById("renderFootnote").checked =
      renderConfig.renderFootnote;
    document.getElementById("highlightTheme").value =
      renderConfig.highlightTheme;
    document.getElementById("renderFront").checked = renderConfig.renderFront;
  }

  loadCodeHighlightTheme(renderConfig.highlightTheme);
}

// 保存设置
function saveSettings() {
  localStorage.setItem("markdownRenderSettings", JSON.stringify(renderConfig));
}

// 修改事件监听器
function setupToggle(id, property) {
  document.getElementById(id).addEventListener("change", function (e) {
    renderConfig[property] = e.target.checked;
    saveSettings();

    renderAll();
  });
}

function loadCodeHighlightTheme(theme) {
  // 移除旧的主题
  const oldLink = document.querySelector("link[data-highlight-theme]");
  if (oldLink) oldLink.remove();

  // 添加新主题
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = `../highlight/highlight/styles/${theme}.min.css`;
  link.dataset.highlightTheme = true;
  document.head.appendChild(link);
}

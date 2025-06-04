---@diagnostic disable: undefined-global
pandoc_crossref.i18n('zh_CN')
metadata {
    ["lang"] = "zh-Hans",

    ["subfigGrid"] = true,
    ["linkReferences"] = true,
    ["codeBlockCaptions"] = true,

    ["chapters"] = true,
    ["chaptersDepth"] = 1,
    ["numberSections"] = true,
    ["sectionsDepth"] = -1,
    ["toc-depth"] = 3,

    ["toc"] = true,
    ["lof"] = true,
    ["lot"] = true,

    ["abstract-section"] = {
        ["section-identifiers"] = {
            "abstract_zh",
            "abstract",
        },
    },

    ["abstract_zh-title"] = "摘要",
    ["abstract-title"] = "Abstract",
}

---@diagnostic disable-next-line: lowercase-global
para = custom_style 'Paragraph'

---@diagnostic disable-next-line: lowercase-global
para_keywords = custom_style 'Keywords'

---@diagnostic disable-next-line: lowercase-global
para_lat = custom_style 'Latin Paragraph'

-- 处理文献引用（[@key]）
function Cite(elem)
    print("Detected Cite: " .. pandoc.utils.stringify(elem))
    return pandoc.Span(elem.content, pandoc.Attr("", {}, {["custom-style"] = "footnote reference"}))
end

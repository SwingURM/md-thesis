# 华东理工大学本科毕业论文markdown写作工具

## 工具

pandoc, pandoc-crossref, python-docx

```bash
pandoc demo.md -o output.docx --filter pandoc-crossref --reference-doc reference.docx --citeproc --csl GB-T-7714—2015（顺序编码，双语，姓名不大写，无URL、DOI，引注有页码）.csl --bibliography cppref.bib
python header.py
```

## 项目文件介绍

reference.docx 控制大部分段落样式。你可以修改解压后的内容，重新打包为.docx来修改。

header.py 使用python-docx进一步控制文档格式。

demo.md 论文内容

cppref.bib 参考文献文件

GB-T....csl 来自Zotero中文社区，仅作少量修改

## Known Issues⚠️

- ❌不能正确处理中文文献条目，因为使用的CSL中涉及了CSL-M语法，这在当前pandoc是不支持的。

- ❌一次性引用多条文献时，格式与要求不符。此外，也不能生成范围（如[2-4]）的引用。

- ❌致谢部分姓名和日期没做。
  
- ⚠️公式的序号标注和引用标注仍需要你手动修正。

- ⚠️参考文献格式与模板轻微不一致，已发现的是空格的显示长度（need help😖）

## TODO
发现word段落设置里有些提到断页啥的，或许也可以利用下。





如果你发现任何我没注意到的问题，可以帮忙留个issue啥的？谢谢




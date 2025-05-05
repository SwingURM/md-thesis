## 工具

pandoc, pandoc-crossref, python-docx

```bash
pandoc demo.md -o output.docx --filter pandoc-crossref --reference-doc reference.docx --citeproc --csl GB-T-7714—2015（顺序编码，双语，姓名不大写，无URL、DOI，引注有页码）.csl --bibliography ref.bib --lua-filter .\cite.lua
python header.py
```

## 项目文件介绍

reference.docx 控制大部分段落样式。你可以修改解压后的内容，重新打包为.docx来修改。

cite.lua 设置脚注引用的字体

header.py 设置表格格式，添加页眉

demo.md 论文内容

GB-T....csl 来自Zotero中文社区，仅作少量修改

## Known Issues⚠

- 不能正确处理中文文献条目，因为使用的CSL中涉及了CSL-M语法，这在当前pandoc是不支持的。

- 似乎不能处理一次引用多条文献从而显示成一个范围^[2-15]的情况。

- 图的命名格式应为`x-y`。

- 参考文献格式与模板不一致，已发现的是空格的显示长度（need help😖）



## TODO

封面，目录，页眉...



如果你发现任何我没注意到的问题，请帮忙提醒我下，谢谢




# 华东理工大学本科毕业论文markdown写作工具

## 工具

pandoc, pandoc-crossref, python-docx

## 使用pixi管理
```bash
pixi run --environment md2doc python header.py
```

## 使用conda管理
```bash
conda env create -f environment.yml
conda activate md2doc
python header.py
```

## 自己管理
自己操作吧。

## 项目文件介绍

/reference 控制大部分段落样式。基于pandoc预生成的reference.docx解压得到，根据毕业论文要求做了修改

header.py 重新生成reference.docx用于指导大部分段落样式。先调用pandoc，再使用python-docx进一步控制文档格式。

demo.md 论文内容

cppref.bib 参考文献文件

GB-T....csl 来自Zotero中文社区，仅作少量修改

## Known Issues⚠️

- ❌不能正确处理中文文献条目，因为使用的CSL中涉及了CSL-M语法，这在当前pandoc是不支持的。

- ❌一次性引用多条文献时，格式与要求不符。此外，也不能生成范围（如[2-4]）的引用。

- ❌致谢部分姓名和日期没做。
  
- ⚠️参考文献格式与模板轻微不一致，已发现的是空格的显示长度（need help😖）

## TODO
没想好要不要为表格等一些类型的段落启用段中分页这种功能。





如果你发现任何我没注意到的问题，可以帮忙留个issue啥的？谢谢




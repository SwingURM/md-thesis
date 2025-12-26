# 华东理工大学本科毕业论文markdown写作工具
本工具完成了毕业论文的大部分docx排版工作，你只需在markdown完成内容的编写而不用时刻操心格式。
本项目含大量ai生成代码，请谨慎使用。


## 用到的工具

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
自行准备所需的软件包。

## 项目文件介绍

/reference 控制大部分段落样式。基于pandoc预生成的reference.docx解压得到，根据毕业论文要求做了修改

header.py 重新生成reference.docx用于指导大部分段落样式。先调用pandoc，再使用python-docx进一步控制文档格式。

demo.md 论文内容

cppref.bib 参考文献文件

GB-T....csl 来自Zotero中文社区，仅作少量修改

## Known Issues⚠️

- ⚠️不能正确处理中文文献条目，因为使用的CSL中涉及了CSL-M语法，这在当前pandoc是不支持的。比较显眼的情况是在英文文献条目中用“等”而不是"et al."，我已经尝试处理了该情况。不了解是不是还有别的情况。

- ⚠️一次性引用多条文献时，格式与要求不符（应为[1][2]而不是[1],[2]）。此外，也不会主动生成范围（如[2-4]）的引用。

- ⚠️致谢部分姓名和日期没做。

- ⚠️（尽管已经强制要求刷新）有时目录页码不正确，请自行刷新。

- ⚠️正文页眉处需要你自己加上适量空格。
  
~~- ⚠️参考文献格式与模板轻微不一致，已发现的是空格的显示长度（need help😖）~~

## TODO
没想好要不要为表格等一些类型的段落启用段中分页这种功能。





如果你发现任何我没注意到的问题，可以帮忙留个issue啥的？谢谢




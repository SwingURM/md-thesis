# 华东理工大学本科毕业论文 markdown 写作工具

本工具完成了毕业论文的大部分 docx 排版工作，你只需在 markdown 完成内容的编写而不用时刻操心格式。
本项目含大量 ai 生成代码，请谨慎使用。
现已尝试支持开题报告。

## 用到的工具

pandoc, pandoc-crossref, python-docx

## 使用 pixi 管理

```bash
pixi run --environment md2doc python thesis.py #毕业论文
pixi run --environment md2doc python open.py #开题报告
```

## 使用 conda 管理

```bash
conda env create -f environment.yml
conda activate md2doc
python thesis.py #毕业论文
python open.py #开题报告
```

## 自己管理

自行准备所需的软件包。

## 项目文件介绍

/reference 控制大部分段落样式。基于 pandoc 预生成的 reference.docx 解压得到，根据毕业论文要求做了修改

header.py, open.py, thesis.py 重新生成 reference.docx 用于指导大部分段落样式。先调用 pandoc，再使用 python-docx 进一步控制文档格式。

demo.md 论文内容

open.md 开题报告内容

cppref.bib 参考文献文件

GB-T....csl 来自 Zotero 中文社区，仅作少量修改

## Known Issues⚠️

- ⚠️ 不能正确处理中文文献条目，因为使用的 CSL 中涉及了 CSL-M 语法，这在当前 pandoc 是不支持的。比较显眼的情况是在英文文献条目中用“等”而不是"et al."，我已经尝试处理了该情况。不了解是不是还有别的情况。

- ⚠️ 一次性引用多条文献时，格式与要求不符（应为[1][2]而不是[1],[2]）。此外，也不会主动生成范围（如[2-4]）的引用。

- ⚠️ 致谢部分姓名和日期没做。

- ⚠️（尽管已经强制要求刷新）有时目录页码不正确，请自行刷新。

- ⚠️ 正文页眉处需要你自己加上适量空格。

~~- ⚠️ 参考文献格式与模板轻微不一致，已发现的是空格的显示长度（need help😖）~~

- ⚠️ 似乎页眉长得不太一样

## TODO

没想好要不要为表格等一些类型的段落启用段中分页这种功能。

如果你发现任何我没注意到的问题，可以帮忙留个 issue 啥的？谢谢

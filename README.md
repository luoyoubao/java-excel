基于Apache POI的JAVA Excel读写工具包
========================
## Introduction

java-excel是基于Apache POI的JAVA Excel读写工具。
写这个工具也是缘由一个项目的Excel读写需求，当我找遍了所有资料，发现没有一个通用的且比较高效的工具可以直接拿来使用，所以在这个项目结束后，对项目中所遇到的问题进行了总结，整理优化了相关代码，供大家学习交流，希望大家多提建议，如果对大家有帮助，我会努力维护好这个工具~


## Features

* 基于Apache POI，主要用于提供JAVA Excel读写功能；
* 支持xls和xlsx，xlsx是通过XML方式进行读写，效率有大幅提升，有效防止写Excel时内存溢出；
* 对POI进行了封装，支持自定义表头、标题、列及尾部的字体、颜色、对齐方式；
* 支持二级标题，支持数据行的纵向合并；
* 支持XLSX格式的大数据写入(数据量较大时建议使用xlsx格式进行写入)；
* 支持设置单元格边框(包括对合并单元格的边框设置)；

当前仅支持这么多功能，如果对大家有用，后面将逐步完善其它功能，如果对你毫无帮助，请勿喷~

##  Invoke

    ExcelWriter writer = ExcelWriterFactory.getWriter(ExcelType.XLSX);
    writer.process(metadata, datas, "/Users/Robert/Desktop/QA_test/expense.xlsx");

如果在使用过程中有任何问题，欢迎提建议、吐槽~

email: luoyoub@163.com
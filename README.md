<!--
author: wangzugang
date: 2016-02-08
title: 基于xml的Excel数据自动转换实现
tags: exml、Excel
category:作品
status: publish
summary:一个基于xml的Excel文件转换器
-->


1.功能说明

该软件可以通过自定义exml脚本实现Excel表格数据的转换。

2.使用环境

该软件目前仅支持windows系统，运行本软件之前，请确保你的电脑安装了微软的office的软件。

3.使用

打开exml.exe，加载写好的exml脚本确定即可。

4.exml语法说明

>1）、exml是一种使用xml基本语法格式，并在 xml语法基础之上实现自定义扩充语法的标签语言。

>2）、exml标签说明：
>>1.文件头，固定格式内容：“<?xml version="1.0" encoding="gb2312"?>”		
>>2.excel为根节点标签，其它标签必须包含在该标签内部。		
>>3.import为导入excel标签，导入excel定义须书写在该标签内部，其name属性值为导入excel文件名，导入模板中只需按序书写excel表格中的列标题即可。		
>>4.export为导出excel标签，导出excel定义及导出数据来源须书写在该标签内部，其name属性值为导出excel文件名。		
>>5.col为列定义标签，该标签在import与export标签中表现形式不同，name属性值为列标题。		
>>6.在导出excel标签中，col标签含有type属性和子节点。type有三个值，empty表示改列为空，index表示改列显示序号，build表示改列的值需要用户自定义，其自定义形式使用子节点来表示。		
>>7.append为导出excel标签中col列标签的子节点，该标签数据类型有两种，string表示是一个单字符串，array表示是一个字符串列表，其中字符串列表是使用split属性值进行分割的（string类型没有split属性）。append标签的数据来源分为两种，如果有value属性则表示数据源为用户自定义数据，数据内容为value属性值；否则数据源为导入表的数据，如果数据源为导入表数据，则该标签必须有table和field属性，其中table指示数据来自哪个数据表（也就是导入的excel表的名称），field属性指示数据来自table属性指示表的哪一个列。ifend为附加属性，只适用于用户自定义的字符串数据，表示如果该值添加在列表中，列表中最后一项是否也添加该字符串，0表示不添加，其它值表示添加。

5.示例

具体书写可参考“测试1.xml文件”。

6.提示

导入文件必须跟xml放在同一目录下。

7.源码分享地址：https://github.com/wzugang/exml




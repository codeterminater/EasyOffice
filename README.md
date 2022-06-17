[原创作者](https://github.com/holdengong/EasyOffice)

# 说明
因业务需要，在原有项目基础上实现了对Word表格导出复杂表格文档的支持，既Word中即有段落数据替换，又有表格数据模板替换时，原有功能不能很好满足自己，故自己进行了封装支持
可满足有表头表格列表数据的替换。该功能封装在CreateComplexWordAsync，具体在Test中进行了示例说明

==IMPORTANT==
原项目核心依赖dotnetcore.npoi,在npoi官方项目中作者已经对dotnetcore.npoi进行了痛斥，在项目调试时也发现基于dotnetcore.npoi的项目无法断点调试获取调试信息，故将依赖更改为npoi













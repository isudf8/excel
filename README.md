# excel导出json并提供给typescript使用，带类型提示
# by @yuanxiao
# 基于@lanmianbao 的ts版配置文件读取模式制作
# environment  node v10.6.0
# 特别适用于前后端都用typescript开发的项目，前后端同时维护一份表，并舍弃各端不需要的数据
# 表格配置需求:
    第一行:描述
    第二行:字段类型(string;number;boolean;自定义Object)
    第三行:前端使用的数据索引
    第四行:后端使用的数据索引
    第五行&&之后:数据内容

    表单开头带 【!】 的表示这是一个备注表，不会被解析
    数据索引一栏 （前端第3行，后端第4行）
        a:带 【*】 的表示以此字段为key，带 【**】 的表示以此字段为key，且是一个列表
        b:索引中带【.】的代表切割符，表示需要以多字段取值，例如  rank.id  需要以 rank 和 id 取值
# 使用方法:
    >node index.js
    >输入需要导出前端表(c)或后端表(s)

# 备注:
    typescript项目需要以来@type/node包才能使用require
# 解析配置表-环境安装
    >
# e7
Epic7EquipmentEvaluateEmulatorExtraExtendedExtreme
#快速上手
- 克隆本仓库（注意submodule的正确克隆）
- 更新数据
- 更新配装方案
- 运行v2.py
#依赖
- openpyxl
# 更新数据
## 标准数据
编辑Data/all.xslx文件，录入英雄、装备、神器
## 非标准数据
编辑new.json文件，录入数据库（Database）中未包含的英雄数据
# 更新配装方案
## 计算的方案
所有位于Plans目录下（不包括次级目录）的.xlsx文件都会被配装器导入进行计算
# 计算
运行v2.py，等待计算完成
# 结果
计算完成后，在Result目录下会出现以方案文件的文件名命名的目录，其中会出现以方案的表名命名的.txt文件，计算结果记录在此文件中
# 贴士
## 方案的存档
可以将本次计算不需要采用的方案存在Plans目录的次级目录下（例如stack）
这些方案都不会被加入此次计算
在需要时则可以方便地拿出来
## 计算的继续
在执行计算时，如果对应的结果文件已经存在，那么对于结果文件中已经计算完成的英雄，会直接采用该结果
## 方案的微调
基于上述计算的可继续性，可以在已经计算的结果中删除某些英雄的内容，然后进行计算，实现部分英雄的更新

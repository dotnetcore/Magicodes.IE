
## 目前这个项目处于早期的需求收集和演示阶段，还不成熟，不建议用到生产环境，请自行评估。

 `Magicodes.IE.Stash` 名字灵感来源于 `Logstash`，向它致敬。

本模块期望为IE用户带来基于规则引擎（脚本）的灵活数据 `转换` 服务。


目前仅支持动态数据导入，欢迎大家提出新的需求和想法。

## 导入模块说明

目前导入模块仅支持Excel文件数据源，后续将会支持更多的数据源，如：csv、数据库表、http接口等。

### 从Excel文件导入:

 导入时我们需要准备两个东西：
 - 数据源文件 （要导入的原始数据）。
 - Dto映射规则。

 映射规则用于描述数据源中的每一列,如何映射到Dto中的属性，在规则里可以通过 `转换器` 管道对数据流进行多次转换，最终输出为需要的值。

 `转换器` 是一段带返回的C#脚本，可以调用宿主项目中已定义的库和方法，为导入提供更多的灵活性。

 #### 导入流程：
  - 实例化引擎： ``` var _excelImporter = new Magicodes.IE.Stash.Import.ExcelImporter(); ```
  - 读取映射规则： ``` var _mappingRules = _excelImporter.LoadDefinitionFromExcelFile("映射规则文件路径"); ```
  - 编译映射规则：```  _excelImporter.Build(); ``` 
    // 编译时，会解析规则中的脚本，生成执行上下文，同时规则中定义的变量也会计算产生值。
  - 读取数据源： var _data = _excelImporter.Resolve("数据源文件路径");

  `_data` 即为导入的数据，类型为 `List<object>`， 集合中项的实际类型为规则中指定的Dto类型。

### 编写映射规则：
映射规则可以放到单独文件，也可以放到数据源一起 ```_excelImporter.LoadDefinitionFromExcelFile()``` 方法提供两个重载，可以明确指定规则定义的Sheet名称，如果不指定，则尝试从默认的名为`$definition$`的工作表中获取规则。

映射规则大概有这些东西：
- 引入的命名空间：将指定的命名空间引入`using`到脚本解析环境中，如：`Magicodes.IE.Stash.Import; Magicodes.IE.StashTests.Extensions`。
- 目标Dto类型：数据源中的每一行将输出为这个类型的实例，最好使用带命名空间的全名，如：`MyProject.Dtos.StudentDto`。
- 变量定义： 可以在映射规则中定义全局变量，变量的值可以是常量，也可以是带返回的脚本，变量可以被规则中其它地方的脚本引用。

具体规则定义，请参考单元测试文件： `Magicodes.IE.StashTests\_res\正确的定义.xlsx`。

## TODO:
- [ ] 重构项目，规范接口和变量的命名，提取抽象接口。
- [ ] 可免Dto导入，直接返回原始数据集合。
- [ ] 支持从csv文件导入。
- [ ] 支持从数据库表导入。
- [ ] 支持从http接口导入。
- [ ] 支持依赖注入，脚本可以直接从DI中请求需要的服务。
- [ ] 增强安全性。
- [ ] 动态导出。


# AUTOSARCP CODE

## 说明
本仓库主要是自学自用自写的AUTOSAR CP code，不涉及任何商业服务


文件结构说明
```bash
.
├── LICENSE
├── README.md
├── doc                      # 文档位置，工程的说明与模块说明
├── lib                      # 引用库位置，如三方静态库
├── script                   # 脚本位置，如编译脚本，CICT脚本
├── src                      # 工程代码位置
│   ├── app                  # app 代码与相关文件
│   ├── appwrp               # app与bsw 交互的集成文件
│   ├── bsw                  # bsw 手写代码
│   ├── bswlib               # bsw 三方代码，如EB/Vector的代码包源码（提取过的代码）
│   └── conf                 # 手写配置代码与工程生成代码
├── test                     # 测试代码，静态与动态测试代码
└── tools                    # 工具工程，如 EB工程，vector工程，CanOE工程，lauterbach工程...等及一些开发工具

```

## 架构
整体架构原则上与AUTOSAR标准保持一致

软件集成架构

![Integration arch](./doc/image/Integration_arch.png)
软件集成将Bsw分成几大类给App提供接口封装：
- IoWrp
- EepWrp
- ComWrp
- NmWrp
- WakeWrp
- DiagWrp


IOHWAB 架构
![IOHWAB ARCH](./doc/image/IoHwab_Arch.png)


## 编译
本仓库以cmake为主，博主makefile语法有待提到

## 文档


## 脚本
目的为了减少工作量，解耦等
如果脚本带来工作量增加，请扔掉它

## 工具
目的为了减少工作量，解耦等
如果工具带来工作量增加，请扔掉它


## 测试








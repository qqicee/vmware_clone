# 使用方法

## 1：安装PowerCli环境。

https://code.vmware.com/web/tool/12.0.0/vmware-powercli

按照实际环境，选择合适的版本，本脚本在

Powercli 11 + VCSA6.7U3

PowerCLI 6.5+VCSA6.5 测试通过。



## 2：编辑任务列表

以work.xlsx为模板，编辑相应的Esxi、模板、虚拟机名称，数据存储，如需要更改CPU、内存、硬盘，可以按需填写，无需更改则留空。



## 3：运行命令

./clone.ps1 -filename work.xlsx -vchost 你的vc_ip

按提示输入用户名密码，开始克隆。





# 注意事项

- VC不能存在一个叫temp_spec的虚拟机自定义配置文件。
- 只适用于仅有一个网卡的Linux主机。不符合条件的，无法更改IP、主机名。
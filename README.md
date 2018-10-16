# ICPCAutoSeats

为ICPC区域赛主办方提供一个一键生成座位号的脚本

# 原理

假设队伍总数为`tot`每次暴力为所有队伍随机分配一个`[1, tot]`间的数字（保证唯一），对应一个场地内的座位号，然后对于同一学校里的任意两队，`check`这两队的曼哈顿距离，如果大于`最小可接受的曼哈顿距离`就重新随机一次，直到满足要求。

# 可修改的配置文件

```python
width  # 场地的宽度（单位：队）
length  # 场地的长度（单位：队）
tot  # 总队伍数
minimal_accepted_distance  # 最小可接受的曼哈顿距离
path  # 源表格的路径，如果你不知道这是啥的话请不要修改，对应的替代方法是把源表格与py文件放在同一目录下并重命名为xlsx.xlsx。
```

# 用法

在`xlxs.xlxs`中放入学校信息，其中：
1. 请保证第一行中没有学校信息（因为考虑到会写表头，所以程序会跳过第一行）。
2. 请预留出第一列，程序会将每个学校的座位号写至第一列。
3. 请将学校写在第二列，程序会根据此来分配座位号。

然后执行python文件：

```
python3 main.py
```

得到一个`new.xlxs`，其中就包含座位号。

# 其它

不知道为什么生成的`new.xlxs`中会包含一些奇奇怪怪的样式，比如说底色和边框线之类的，大家自行手动更改吧。

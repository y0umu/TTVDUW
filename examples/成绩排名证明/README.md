`2022级智能建造学生成绩排名_datafeed.xlsx`里面学生的学号、姓名完全是虚构的。如有雷同纯属巧合。

姓名的虚构使用了[Faker](https://faker.readthedocs.io/en/master/)。如果你一定要知道的话：
```python
from faker import Faker
fake = Faker('zh_CN')
for i in range(20):
    print(fake.name())
```
import pandas as pd

PR = ["攻击","穿甲","溅伤", "吸血",
      "免伤","反伤","气血","防御", "回血",
      "速度","招架","命中","闪避",
      "真元","聚气","破招","破闪",
      "拳法","剑法","刀法","棍法",
    ]

EPR = ["真元","攻击","臂力",
       "免伤","身法","气血","防御",
    ]                  

EQ = ["紫薇护腕","乌斯护腕",
      "玉衡护腰","吉达护腰",
      "乌蚕宝甲","呼如木甲",
      "青云戒","苏布宝戒",
      "青云坠","其木格坠",
      "破军指套","万仞重剑","碎月重刀","贪狼棍","碎虚枪",
      ]

class Property():
    """装备属性"""
    def __init__(self, pr_type = PR[0] , value=0.0):
        self.pr_type = pr_type
        self.value = value

    def __str__(self):
        return f'{self.pr_type}:{self.value}'

class ExtraProperty():
    """装备附加属性"""
    def __init__(self, epr_type = EPR[0] , value=0.0):
        self.epr_type = epr_type
        self.value = value

    def __str__(self):
        return f'{self.epr_type}:{self.value}'

class Equipment():
    def __init__(self, path:str , sheet_name:str , eq_type:str , extra = ExtraProperty(), remark = ''):
        """初始化\n
        必须有文件路径;\n
        必须有列表号（装备名）;\n
        可选附加;\n
        可选备注（可记录序列号）\n"""
        if path == '':
            raise ValueError(f'缺少路径')
        if sheet_name == '':
            raise ValueError(f'缺少表名')
        self.df = pd.read_excel(path,sheet_name=sheet_name,index_col=0)
        
        if eq_type not in EQ:
            raise ValueError(f'并非合理的装备')
        else:
            self.eq_type = eq_type
            
        self.extra = extra
        self.remark = remark

    def trans(self,pr:Property):
        """专有属性转换"""
        return Property(pr.pr_type,pr.value * 20) if pr.pr_type in PR[-4:] else pr

    def calculate(self):
        """计算精炼属性（列）"""
        for col in self.df.columns:
            if col == 0:        # 第一列为初始值不计算
                continue
            for row in self.df.index:
                if row in ['1','2','3']:   # 仅看三行精炼数值部分
                    pr_type = self.df.loc[row,col].pr_type
                    value = self.df.loc[row,col].value
                    if value != 0:
                        if col <= 20:
                            value = round (value / 0.6 , 1)
                        elif col <= 40:
                            value = round (value / 0.7 , 1)
                        elif col <= 60:
                            value = round (value / 0.8, 1)
                        elif col <= 80:
                            value = round (value / 0.9, 1)
                        else:
                            value = round (value / 1.0, 1)
                        self.df.loc['精炼',col] = Property(pr_type,value)
    
    def __str__(self):
        return self.df.to_string()+'\n\n'+str(self.eq_type)+'\n'+self.remark+'\n'+'附加'+str(self.extra)+'\n'

def main():
    eqs = {}
    eqs['path'] = 'test.xlsx'


    eqs['sheet_name'] = 'eq1'
    eqs['eq_type'] = EQ[-1]
    eqs['extra'] = ExtraProperty(EPR[0],10)
    eqs['remark'] = '第12件'
    eqs['sheet_name'] = eqs['remark']+eqs['eq_type']

    w = Equipment(**eqs)
    print(w)



if __name__ == '__main__':
    main()

# df = pd.DataFrame(index=['1','2','3','精炼'])
# id = -1
# default = 0
# df [default] = [Property(PR[0],100),Property(PR[id],5.6),Property(PR[0],55),Property()]
# times = 20
# df [times] = [Property(),Property(PR[id],3.4),Property(),Property()]

# eq = Equipment(eq_type = EQ[-1], df = df,extra = ExtraProperty(EPR[0],10),remark= '第12件')
# eq.calculate()
# print(eq)


# # with pd.ExcelWriter('test.xlsx',engine='openpyxl',mode='w') as writer:
# #     eq.df.to_excel(writer,sheet_name='eq1')


# # eq0 = pd.read_excel('test.xlsx',sheet_name='eq1',index_col=0)

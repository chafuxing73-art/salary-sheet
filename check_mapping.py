import pandas as pd

df = pd.read_excel('总部人员信息表.xlsx', header=None)
mapping = {}
for _, row in df.iterrows():
    if len(row) >= 2 and pd.notna(row[0]) and pd.notna(row[1]):
        name = str(row[0]).strip()
        if '-' in name:
            name_parts = name.split('-')
            if len(name_parts) > 1:
                name = name_parts[-1].strip()
        userid = str(row[1]).strip()
        mapping[name] = userid

print('映射表中的姓名（前20个）:')
print(list(mapping.keys())[:20])
print(f'\n总共有 {len(mapping)} 个映射')

salary_names = ['查甫兴', '查', '1兴', '查兴', '22甫兴', '12兴', '22兴', '23兴', '许晓燕']
print('\n工资表中的姓名及其匹配情况:')
for name in salary_names:
    if name in mapping:
        print(f'  ✅ {name} -> {mapping[name]}')
    else:
        print(f'  ❌ {name} 未找到映射')
# -*- coding: utf-8 -*-
#
# @Time    : 2022-03-03 16:50:46
# @Author  : Yunlong Shi
# @Email   : syljoy@163.com
# @FileName: op_excel.py
# @Software: PyCharm
# @Github  : https://github.com/syljoy
# @Desc    : 题库数据加载

import pandas as pd


def get_qb(path=r'./QB/初赛题库1（选拔赛题库）20220217.xlsx'):
    single_qb = pd.read_excel(path, sheet_name='单选题')
    multiple_qb = pd.read_excel(path, sheet_name='多选题')
    tf_qb = pd.read_excel(path, sheet_name='判断题')

    single_qb = single_qb.drop(['* 年级', '* 学科', '知识点', '* 解析', '* 分数', '* 难度', '导入结果'], axis=1)  # 删除列
    multiple_qb = multiple_qb.drop(['* 年级', '* 学科', '知识点', '* 解析', '* 分数', '* 难度'], axis=1)  # 删除列
    tf_qb = tf_qb.drop(['* 年级', '* 学科', '知识点', '* 解析', '* 分数', '* 难度', '导入结果'], axis=1)  # 删除列

    return single_qb, multiple_qb, tf_qb


if __name__ == '__main__':
    excel_path = r'./QB/初赛题库1（选拔赛题库）20220217.xlsx'

    single_qb, multiple_qb, tf_qb = get_qb(excel_path)
    print(single_qb.head(2))
    print(multiple_qb.head(2))
    print(tf_qb.head(2))

    print('*-' * 20)
    s = '于坚持和完善中国特色社会主义制度、推进国家治理体系和治理能力现代化若干重大问题的决定》指出，要提高防范抵御国家安全风险能力'
    s = '能力'
    print(len(multiple_qb.loc[multiple_qb['* 题干'].str.contains(s)]))
    print(multiple_qb.loc[multiple_qb['* 题干'].str.contains(s)]['* 标答'])
    # print(data)

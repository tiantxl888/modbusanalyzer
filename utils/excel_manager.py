#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

class ExcelManager:
    """
    统一从一个 Excel 文件的不同 sheet 中加载所有配置与参数：
      - ConnectionConfig: 通信配置（来自 connectionConfig.txt）
      - LocalSettings:   本地设置（来自 localSettings.txt）
      - Language:        界面语言字典（来自 language.json）
      - Parameters:      寄存器参数表（来自 parameter.xlsx）
    """

    def __init__(self, filepath: str):
        """
        :param filepath: 合并后的配置文件路径，如 'config_and_params.xlsx'
        """
        self.filepath = filepath

    def load_connection(self, sheet_name: str = 'ConnectionConfig') -> dict:
        """
        加载通信配置，返回 {Key: Value} 字典
        """
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, dtype=str)
        # 部分值可能为空，用空字符串代替
        df['Value'] = df['Value'].fillna('')
        return dict(zip(df['Key'], df['Value']))

    def load_local_settings(self, sheet_name: str = 'LocalSettings') -> dict:
        """
        加载本地设置，尝试将数字和布尔值转为对应类型
        """
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, dtype=str)
        df['Value'] = df['Value'].fillna('')
        settings = {}
        for key, raw in zip(df['Key'], df['Value']):
            val = raw.strip()
            # 布尔
            if val.lower() in ('true', 'false'):
                settings[key] = (val.lower() == 'true')
            # 整数
            elif val.isdigit():
                settings[key] = int(val)
            # 浮点
            else:
                try:
                    settings[key] = float(val)
                except ValueError:
                    settings[key] = val
        return settings

    def load_language(self, lang: str = 'zh', sheet_name: str = 'Language') -> dict:
        """
        加载指定语言的界面文案，返回 {Key: 文案} 字典
        """
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, dtype=str)
        if lang not in df.columns:
            raise ValueError(f"Language sheet 中缺少列: {lang}")
        return dict(zip(df['Key'], df[lang].fillna('')))

    def load_parameters(self, sheet_name: str = 'Parameters') -> pd.DataFrame:
        """
        加载寄存器参数表，返回 pandas.DataFrame
        """
        df = pd.read_excel(self.filepath, sheet_name=sheet_name, header=1)
        
        # 确保dataType列的处理
        if 'dataType' in df.columns:
            # 修复NaN值和空字符串
            df['dataType'] = df['dataType'].apply(
                lambda x: 'UNSIGNED' if pd.isna(x) or str(x).upper() == 'NAN' or str(x).strip() == '' else str(x).strip().upper()
            )
            
            # 特殊处理已知需要SIGNED的地址
            known_signed_addrs = ['10000', '10001', '10002', '10003', '10004', '10005', '10006', '10007', '10008', '10009', '10010', '10011', '10012']  # 已知需要SIGNED类型的地址列表
            for addr in known_signed_addrs:
                if addr in df['addr'].astype(str).values:
                    idx = df[df['addr'].astype(str) == addr].index
                    if len(idx) > 0:
                        df.loc[idx, 'dataType'] = 'SIGNED'
                        print(f"DEBUG: Excel加载时修复地址{addr}的数据类型为SIGNED")
        
        return df

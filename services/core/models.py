#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
データモデル定義モジュール

システム全体で使用する共通のデータクラスを定義
"""

from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime


@dataclass
class DeliveryItem:
    """配送アイテムのデータクラス"""
    item_code: str = ""
    item_name: str = ""
    quantity: int = 0
    unit: str = ""
    unit_price: float = 0.0
    total_price: float = 0.0
    delivery_date: str = ""
    warehouse: str = ""
    notes: str = ""
    
    def __post_init__(self):
        """初期化後の処理"""
        # 合計金額を自動計算
        if self.unit_price > 0 and self.quantity > 0:
            self.total_price = self.unit_price * self.quantity


@dataclass
class DeliveryDocument:
    """配送文書のデータクラス"""
    document_id: str = ""
    document_date: str = ""
    supplier: str = ""
    delivery_destination: str = ""
    total_amount: float = 0.0
    status: str = "pending"  # pending, validated, error
    items: List[DeliveryItem] = None
    notes: str = ""
    created_at: str = ""
    
    def __post_init__(self):
        """初期化後の処理"""
        if self.items is None:
            self.items = []
        
        if not self.created_at:
            self.created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # 合計金額を自動計算
        self._calculate_total_amount()
    
    def _calculate_total_amount(self):
        """合計金額を計算"""
        if self.items:
            self.total_amount = sum(item.total_price for item in self.items)
    
    def add_item(self, item: DeliveryItem):
        """アイテムを追加"""
        self.items.append(item)
        self._calculate_total_amount()
    
    def get_item_count(self) -> int:
        """アイテム数を取得"""
        return len(self.items)
    
    def get_total_quantity(self) -> int:
        """総数量を取得"""
        return sum(item.quantity for item in self.items)


@dataclass
class MasterItem:
    """マスタアイテムのデータクラス"""
    item_code: str = ""
    item_name: str = ""
    supplier: str = ""
    category: str = ""
    unit_price: float = 0.0
    delivery_destinations: List[str] = None
    warehouse: str = ""
    notes: str = ""
    
    def __post_init__(self):
        if self.delivery_destinations is None:
            self.delivery_destinations = []


@dataclass
class ValidationError:
    """バリデーションエラーのデータクラス"""
    error_type: str = ""
    item_code: str = ""
    item_name: str = ""
    expected_value: str = ""
    actual_value: str = ""
    description: str = ""
    document_id: str = ""
    created_at: str = ""
    
    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')


@dataclass
class OrderItem:
    """受注アイテムのデータクラス（内部使用用）"""
    item_code: str = ""
    quantity: int = 0
    extracted_date: str = ""
    source_file: str = ""
    
    def __post_init__(self):
        if not self.extracted_date:
            self.extracted_date = datetime.now().strftime('%Y-%m-%d')

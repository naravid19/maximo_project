"""
Utility helper functions for Maximo App.

This module contains reusable helper functions extracted from views.py
to improve code maintainability and reduce duplication (DRY principle).
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import Any, List, Optional


def forward_fill_column(df: pd.DataFrame, column_name: str) -> List[Any]:
    """
    Forward fill values in a DataFrame column, replacing -1 with the previous value.
    
    This function implements a forward-fill pattern commonly used when processing
    Excel data where cells are merged and only contain values in the first row.
    
    Args:
        df: DataFrame containing the column to process.
        column_name: Name of the column to forward fill.
        
    Returns:
        List of values with -1 replaced by the previous non-(-1) value.
        
    Example:
        >>> df = pd.DataFrame({'KKS': ['A', -1, -1, 'B', -1]})
        >>> forward_fill_column(df, 'KKS')
        ['A', 'A', 'A', 'B', 'B']
    """
    result = []
    temp_value = ''
    
    for value in df[column_name]:
        if value != -1:
            temp_value = value
        result.append(temp_value)
    
    return result


def update_comment(df: pd.DataFrame, condition: pd.Series, column_name: str, message: str) -> None:
    """
    Update a comment column by appending a message when condition is True.
    
    Args:
        df: DataFrame to update.
        condition: Boolean Series indicating which rows to update.
        column_name: Name of the comment column.
        message: Message to append to existing comments.
    """
    df.loc[condition, column_name] = df.loc[condition, column_name].apply(
        lambda x: f"{x}, {message}" if x else message
    )


def replace_or_append_comment(
    df: pd.DataFrame,
    condition: pd.Series,
    comment_col: str,
    message: str,
    replace_message: Optional[str] = None
) -> None:
    """
    Replace or append a comment in a DataFrame column.
    
    If replace_message is provided and found in the existing comment, it will be
    replaced with the new message. Otherwise, the message is appended.
    
    Args:
        df: DataFrame to update.
        condition: Boolean Series indicating which rows to update.
        comment_col: Name of the comment column.
        message: New message to add.
        replace_message: Optional message to replace if found.
    """
    def update_comment_text(existing_comment: str) -> str:
        if replace_message and replace_message in str(existing_comment):
            return str(existing_comment).replace(replace_message, message)
        elif existing_comment:
            return f"{existing_comment}, {message}"
        return message
    
    df.loc[condition, comment_col] = df.loc[condition, comment_col].apply(update_comment_text)


def convert_duration(value: Any) -> Any:
    """
    Convert duration value to a number, handling various input formats.
    
    Args:
        value: Input value that may be a number, string, or other type.
        
    Returns:
        Converted numeric value, or original value if conversion fails.
    """
    if pd.isna(value):
        return value
    
    if isinstance(value, (int, float)):
        return value
    
    if isinstance(value, str):
        try:
            # Remove any whitespace and try to convert
            cleaned = value.strip().replace(',', '')
            return float(cleaned)
        except ValueError:
            return value
    
    return value


def is_date(value: Any) -> bool:
    """
    Check if a value can be parsed as a date.
    
    Args:
        value: Value to check.
        
    Returns:
        True if the value can be parsed as a date, False otherwise.
    """
    if pd.isna(value):
        return False
    
    try:
        pd.to_datetime(value)
        return True
    except (ValueError, TypeError):
        return False


def parse_dates(date_series: pd.Series) -> pd.Series:
    """
    Parse a Series of date values to datetime objects.
    
    Args:
        date_series: Series containing date values.
        
    Returns:
        Series with parsed datetime values, NaT for unparseable values.
    """
    return pd.to_datetime(date_series, errors='coerce')


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean and standardize DataFrame column names.
    
    Converts all column names to uppercase and replaces spaces with underscores.
    
    Args:
        df: DataFrame with columns to clean.
        
    Returns:
        DataFrame with cleaned column names.
    """
    df.columns = [
        col.strip().upper().replace(' ', '_') if isinstance(col, str) else col
        for col in df.columns
    ]
    return df


def uppercase_string_column(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    """
    Convert all string values in a column to uppercase.
    
    Args:
        df: DataFrame containing the column.
        column_name: Name of the column to uppercase.
        
    Returns:
        DataFrame with uppercased column values.
    """
    df[column_name] = df[column_name].apply(
        lambda x: x.upper() if isinstance(x, str) else x
    )
    return df


def strip_string_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """
    Strip whitespace from string values in specified columns.
    
    Args:
        df: DataFrame containing the columns.
        columns: List of column names to strip.
        
    Returns:
        DataFrame with stripped column values.
    """
    for col in columns:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda x: x.strip() if isinstance(x, str) else x
            )
    return df


def get_grouping_text(selected_order: List[str]) -> str:
    """
    Convert a list of grouping options to a display text.
    
    Args:
        selected_order: List of grouping option codes.
        
    Returns:
        Comma-separated string of grouping options.
    """
    return ', '.join(selected_order) if selected_order else 'None'

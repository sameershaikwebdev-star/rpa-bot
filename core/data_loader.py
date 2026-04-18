"""
Data Loader Module
Reads Excel / CSV files using Pandas and validates the data
before passing it to the RPA bot engine.
"""

import pandas as pd
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


class DataLoader:
    """
    Loads, validates, and prepares Excel/CSV data for the RPA bot.
    """

    SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}

    def __init__(self, filepath: str):
        self.filepath = Path(filepath)
        self.df: pd.DataFrame = pd.DataFrame()
        self._validate_file()

    # ─── File Validation ──────────────────────────────────────────────────────

    def _validate_file(self):
        if not self.filepath.exists():
            raise FileNotFoundError(f"File not found: {self.filepath}")
        if self.filepath.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Unsupported format '{self.filepath.suffix}'. "
                f"Use one of: {self.SUPPORTED_EXTENSIONS}"
            )

    # ─── Reading ──────────────────────────────────────────────────────────────

    def load(self, sheet_name: str = 0, header_row: int = 0) -> pd.DataFrame:
        """
        Load the file into a DataFrame.
        Supports .xlsx, .xls, and .csv automatically.
        """
        ext = self.filepath.suffix.lower()
        if ext in {".xlsx", ".xls"}:
            self.df = pd.read_excel(
                self.filepath,
                sheet_name=sheet_name,
                header=header_row,
            )
        else:
            self.df = pd.read_csv(self.filepath, header=header_row)

        logger.info(
            f"📂 Loaded '{self.filepath.name}': "
            f"{len(self.df)} rows × {len(self.df.columns)} columns"
        )
        return self.df

    # ─── Cleaning ─────────────────────────────────────────────────────────────

    def clean(self, required_columns: list[str] = None) -> pd.DataFrame:
        """
        • Strip whitespace from string columns
        • Drop fully empty rows
        • Optionally keep only required_columns
        • Fill NaN with empty string for safe form filling
        """
        if self.df.empty:
            raise RuntimeError("DataFrame is empty. Call load() first.")

        # Normalise column names
        self.df.columns = [str(c).strip() for c in self.df.columns]

        # Strip leading/trailing spaces in string columns
        str_cols = self.df.select_dtypes(include="object").columns
        self.df[str_cols] = self.df[str_cols].apply(
            lambda col: col.str.strip()
        )

        # Drop fully empty rows
        before = len(self.df)
        self.df.dropna(how="all", inplace=True)
        dropped = before - len(self.df)
        if dropped:
            logger.info(f"🧹 Dropped {dropped} empty rows")

        # Column filter
        if required_columns:
            missing = set(required_columns) - set(self.df.columns)
            if missing:
                raise ValueError(f"Missing required columns: {missing}")
            self.df = self.df[required_columns]

        # Fill NaN → empty string (safe for Selenium send_keys)
        self.df.fillna("", inplace=True)

        logger.info(f"✅ Clean DataFrame: {len(self.df)} rows ready")
        return self.df

    # ─── Export ───────────────────────────────────────────────────────────────

    def to_records(self) -> list[dict]:
        """Convert DataFrame to a list of dicts for the bot engine."""
        return self.df.to_dict(orient="records")

    # ─── Quick Info ───────────────────────────────────────────────────────────

    def summary(self) -> dict:
        """Return a quick summary of the loaded data."""
        return {
            "file":        str(self.filepath),
            "rows":        len(self.df),
            "columns":     list(self.df.columns),
            "null_counts": self.df.isnull().sum().to_dict(),
        }

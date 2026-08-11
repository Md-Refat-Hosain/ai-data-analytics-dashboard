import pandas as pd
import json


def load_data(file_path_or_buffer):
    """Loads CSV dataset with fallback encoding handling for non-breaking spaces & legacy encodings."""
    encodings_to_try = ["utf-8", "latin1", "iso-8859-1", "cp1252"]

    for encoding in encodings_to_try:
        try:
            # Reset buffer pointer if it's an uploaded Streamlit BytesIO object
            if hasattr(file_path_or_buffer, "seek"):
                file_path_or_buffer.seek(0)

            df = pd.read_csv(file_path_or_buffer, encoding=encoding)

            # Clean up column names (strip spaces)
            df.columns = df.columns.str.strip()
            return df
        except (UnicodeDecodeError, Exception):
            continue

    raise ValueError(
        "Unable to decode the CSV file with standard encodings (utf-8, latin1, iso-8859-1, cp1252)."
    )


def generate_schema_catalog(df: pd.DataFrame) -> str:
    """Generates a JSON schema description for the LLM."""
    schema = []
    for col in df.columns:
        col_info = {
            "name": col,
            "type": str(df[col].dtype),
            "sample_values": df[col].dropna().unique()[:3].tolist(),
        }
        schema.append(col_info)
    return json.dumps(schema, indent=2)


def detect_iqr_outliers(df: pd.DataFrame):
    """Detects outliers in numerical columns using the IQR method."""
    numeric_cols = df.select_dtypes(include=["float64", "int64"]).columns
    outlier_summary = {}

    for col in numeric_cols:
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR

        outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)]
        outlier_summary[col] = {
            "outlier_count": len(outliers),
            "percentage": round((len(outliers) / len(df)) * 100, 2),
            "lower_bound": round(lower_bound, 2),
            "upper_bound": round(upper_bound, 2),
        }
    return outlier_summary

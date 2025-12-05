
import pandas as pd


class GisImport:

    @classmethod
    def import_data(cls, details_file_path: str, data_columns: list[str]):
        df = pd.read_csv(details_file_path)
        df.columns = [col.lower() for col in df.columns]

        data_columns = [data_col.lower() for data_col in data_columns]
        df = df[data_columns]

        return df

    @classmethod
    def import_gis(cls, gis_zip_path: str):
        import geopandas as gpd
        gdf = gpd.read_file(gis_zip_path, encoding='utf-8')
        gdf.columns = [col.lower() for col in gdf.columns]

        return gdf


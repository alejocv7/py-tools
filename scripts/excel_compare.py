"""
Compare two excel .xlsx files
"""

import pathlib

import pandas as pd
import xlsxwriter


class ExcelCompare:
    def __init__(
        self,
        old_file_path: str,
        new_file_path: str,
        index_name: str = "_index_",
        other_ref_cols: list[str] | None = None,
    ) -> None:
        self.old_file_path = pathlib.Path(old_file_path)
        self.new_file_path = pathlib.Path(new_file_path)
        self.out_file_path = rf"{self.new_file_path.parent}/comparison.xlsx"

        self.index_name = index_name
        self.other_ref_cols = other_ref_cols
        self.ignored_cols: list[str] = []

        self.old_excel = pd.read_excel(self.old_file_path, sheet_name=None)
        self.new_excel = pd.read_excel(self.new_file_path, sheet_name=None)

        # self.temp_compare()

    def temp_compare(self) -> None:
        self.old_file = pd.read_excel(self.old_file_path, sheet_name=None)
        self.new_file = pd.read_excel(self.new_file_path, sheet_name=None)

        for sheet_old, sheet_new in zip(self.old_file, self.new_file):
            df_old = self.old_file[sheet_old]
            df_new = self.new_file[sheet_new]
            difference = df_old[df_old != df_new]
            print(difference)

    def get_excel_diff(self) -> dict[str | int, dict[str, pd.DataFrame]]:
        excel_diff: dict[str | int, dict[str, pd.DataFrame]] = {}

        for sheet_name in self.old_excel:
            print(f"** Comparing sheet {sheet_name} **")

            if self.index_name == "_index_":
                self.old_excel[sheet_name].index.set_names(
                    self.index_name, inplace=True
                )
                self.new_excel[sheet_name].index.set_names(
                    self.index_name, inplace=True
                )
            else:
                self.old_excel[sheet_name].set_index(self.index_name, inplace=True)
                self.new_excel[sheet_name].set_index(self.index_name, inplace=True)

            excel_diff[sheet_name] = self.get_sheet_diff(
                self.old_excel[sheet_name], self.new_excel[sheet_name]
            )

            print(f"Done {sheet_name}")

        return excel_diff

    def get_sheet_diff(
        self, df_old: pd.DataFrame, df_new: pd.DataFrame
    ) -> dict[str, pd.DataFrame]:
        self.delete_ignored_and_conflicting_cols(df_old, df_new)

        df_merged = pd.merge(
            df_old, df_new, on=self.index_name, how="outer", indicator=True
        )

        # Take those rows present in both sheets and compare their values
        df_idx_in_both = df_merged.index[df_merged["_merge"] == "both"]
        df_old_both = df_old.loc[df_idx_in_both]
        df_new_both = df_new.loc[df_idx_in_both]
        try:
            df_cell_cmp = df_old_both.compare(
                df_new_both,
                result_names=(self.old_file_path.stem, self.new_file_path.stem),
            ).fillna("")
        except ValueError as e:
            raise ValueError("Possibly duplicate index") from e

        # Add ref columns to the left of the df_cell_cmp
        if self.other_ref_cols and not df_cell_cmp.empty:
            df_ref = df_old_both.loc[df_cell_cmp.index, self.other_ref_cols]
            multi_idx_tup = [(ref_col, "-") for ref_col in self.other_ref_cols]
            df_ref.columns = pd.MultiIndex.from_tuples(multi_idx_tup)
            df_cell_cmp = pd.concat([df_ref, df_cell_cmp], axis=1)

        df_idx_in_old_only = df_merged.index[df_merged["_merge"] == "left_only"]
        df_idx_in_new_only = df_merged.index[df_merged["_merge"] == "right_only"]

        return {
            "changed": df_cell_cmp,
            "added": df_new.loc[df_idx_in_new_only],
            "deleted": df_old.loc[df_idx_in_old_only],
        }

    def write_to_excel(self) -> None:
        print("Writing diff to Excel ...")
        with pd.ExcelWriter(self.out_file_path) as writer:
            for sheet_name, sheet_dict in self.get_excel_diff().items():
                if not all(df.empty for df in sheet_dict.values()):
                    self.sheet_diff_to_excel(sheet_dict, sheet_name, writer)
        print("Finished writing to Excel!")

    def set_ignored_cols(self, cols: list[str]) -> None:
        self.ignored_cols = cols

    def delete_ignored_and_conflicting_cols(
        self, df1: pd.DataFrame, df2: pd.DataFrame
    ) -> None:
        """Delete ignored cols and those that are not common to both dataframes"""
        cols_to_delete = df1.columns.symmetric_difference(df2.columns).union(
            self.ignored_cols
        )
        df2.drop(df2.columns.intersection(cols_to_delete), inplace=True, axis=1)
        df1.drop(df1.columns.intersection(cols_to_delete), inplace=True, axis=1)

    def sheet_diff_to_excel(
        self,
        sheet_dict: dict[str, pd.DataFrame],
        sheet_name: str | int,
        excel_writer: pd.ExcelWriter,
    ) -> None:
        worksheet = excel_writer.book.add_worksheet(sheet_name)
        excel_writer.sheets[worksheet.name] = worksheet

        for key, df_group in sheet_dict.items():
            if not sheet_dict[key].empty:
                self.df_to_excel(df_group, key.title(), excel_writer, worksheet)

    def df_to_excel(
        self,
        df: pd.DataFrame,
        df_title: str,
        excel_writer: pd.ExcelWriter,
        worksheet: xlsxwriter.Workbook.worksheet_class,
    ) -> None:
        worksheet.dim_rowmax = worksheet.dim_rowmax or 0
        worksheet.write_string(worksheet.dim_rowmax, 0, df_title)
        df.to_excel(
            excel_writer, sheet_name=worksheet.name, startrow=worksheet.dim_rowmax + 1
        )
        worksheet.dim_rowmax += 2


excel_cmp = ExcelCompare("scripts/old.xlsx", "scripts/new.xlsx")
excel_cmp.write_to_excel()

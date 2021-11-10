from pathlib import Path
from typing import List

import dynamic_yaml
import pandas as pd

# Directory constants
from data_loaders import BSTLoader, DataConsolidator, DataLoader
from excel import ExcelFormats, ExcelGenerator
from utils import DynamicResolver, reset_index

# TODO make it be able to read from a URL.

DASHBOARD_DIRECTORY = r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards"





class DashboardGenerator:

    def __init__(self, config_path: [str, Path], exclusions_file_path: [str, Path], issue: int):

        with open(config_path) as file:
            self.config = dynamic_yaml.load(file, recursive=True)

        with open(exclusions_file_path) as file:
            self.exclusions = dynamic_yaml.load(file, recursive=True)
        self.ghd_logo_path = Path().cwd().parent / self.config.logos.ghd
        self.client_logo_path = Path().cwd().parent / self.config.logos.client
        self.issue = issue
        self.bst: BSTLoader
        self.previous_dashboard: DataLoader
        self.project_manager_loaders: List[DataLoader] = []
        self.consolidator: DataConsolidator
        self.resolver = DynamicResolver()
        self.margins = self.config.margins

    def _loader(self, data_path: [str, Path], is_pm: bool = False) -> DataLoader:
        loader = DataLoader(
            data_path=data_path,
            exclusions=self.exclusions,
            column_order=self.resolver.resolve(self.config.column.lists.col_order),
            date_cols=self.resolver.resolve(self.config.column.lists.date_cols),
        )
        loader.rename_columns(self.config.mapping.legacy)
        loader.remove_duplicates(duplicate_col=self.config.column.names.project_number)
        loader.exclude()
        loader.add_missing_columns()
        if not is_pm:
            loader.clear_data(self.config.column.lists.mandatory)
        loader.set_column_order()
        return loader

    def _dashboard_maker(self, sheet_name: str, workbook, df: pd.DataFrame, formatter: ExcelFormats) -> ExcelGenerator:
        dashboard = ExcelGenerator(
            workbook=workbook,
            margins=self.margins["inch"],
            issue=self.issue,
            df=df,
            formats=formatter,
            pm_col=self.config.column.names.pm,
            project_col=self.config.column.names.project_number,
            column_widths=self.resolver.resolve(self.config.column.widths),
            ghd_image_path=self.ghd_logo_path,
            client_image_path=self.client_logo_path,
            schedule_col=self.config.column.names.sch,
            data_validation=self.resolver.resolve(self.config.data_validation),
        )
        dashboard.add_worksheet(sheet_name)
        dashboard.setup_worksheet(sheet_name)
        dashboard.setup_header(sheet_name)
        return dashboard

    def load_bst(self, bst_path: [str, Path]):
        self.bst = BSTLoader(
            path=bst_path,
            exclusions=self.exclusions,
            column_order=self.resolver.resolve(self.config.column.lists.col_order),
            date_cols=None,
        )
        self.bst.remove_inactive(activity_col=self.config.column.names.project_status, activity_value=self.config.filters.project_status)
        self.bst.rename_columns(self.config.mapping.bst)
        self.bst.remove_duplicates(duplicate_col=self.config.column.names.project_number)
        self.bst.remove_proposals(
            proposal_col=self.config.column.names.task_code,
            proposal_str=self.config.filters.proposal_code,
            prop_number_lb=self.config.filters.proposal_lower_bound,
            prop_number_ub=self.config.filters.proposal_upper_bound,
        )
        self.bst.select_columns(self.resolver.resolve(self.config.column.lists.bst_cols))
        self.bst.add_missing_columns()
        self.bst.exclude()
        self.bst.set_column_order()

    def load_prev(self, dashboard_path: [str, Path]):
        self.previous_dashboard = self._loader(dashboard_path)

    def load_pm_sheets(self, pm_sheet_folder: [str, Path]):
        pm_sheet_folder = Path(pm_sheet_folder)
        for file_path in pm_sheet_folder.glob("*.xlsx"):
            self.project_manager_loaders.append(self._loader(file_path))

    def create_consolidator(self):
        self.consolidator = DataConsolidator(
            bst_loader=self.bst,
            project_col=self.config.column.names.project_number,
            project_manager_col=self.config.column.names.pm,
            column_order=self.resolver.resolve(self.config.column.lists.col_order),
            project_manager_sheets=self.project_manager_loaders
        )
        self.consolidator.find_new_data()
        self.consolidator.create_master()

    def update_master_with_pm(self):
        self.consolidator.load_pm_sheets_to_master()

    def create_formatter(self, workbook, new_pms, new_projects) -> ExcelFormats:
        formatter = ExcelFormats(
            workbook=workbook,
            base=self.config.formats.base,
            on_hold=self.config.formats.on_hold,
            at_risk=self.config.formats.at_risk,
            behind_schedule=self.config.formats.behind_schedule,
            on_track=self.config.formats.on_track,
            mandatory_input=self.config.formats.mandatory_input,
            new_project=self.config.formats.new_project,
            new_pm=self.config.formats.new_pm,
            client_project_number_error=self.config.formats.client_project_number_error,
            header=self.config.formats.header,
            mandatory_cols=self.resolver.resolve(self.config.column.lists.mandatory),
            new_pms=new_pms,
            new_projects=new_projects,
            client_project_number_errors=[],  # Ignoring for now as I dont think we need any more
        )
        return formatter

    def _create_dashboard(self, out_path: [str, Path], df, sheet_name: str = "Dashboard"):
        writer = pd.ExcelWriter(out_path, engine="xlsxwriter")
        workbook = writer.book
        formatter = self.create_formatter(
            workbook=workbook,
            new_pms=self.consolidator.new_project_managers,
            new_projects=self.consolidator.new_project_managers,
        )

        dashboard = self._dashboard_maker(
            sheet_name=sheet_name,
            workbook=workbook,
            df=df,
            formatter=formatter
        )
        dashboard.write_to_sheet(sheet_name)
        dashboard.add_data_validation(sheet_name)
        workbook.close()

    def create_master_dashboard(self, out_path: [str, Path], sheet_name: str = "Dashboard", ):

        self._create_dashboard(out_path, self.consolidator.master, sheet_name, )

    def get_pm_out_path(self, out_path: [str, Path], name, new_folder: str = None):
        if new_folder:
            out_path = out_path / new_folder / f"{name}.xlsx"
        else:
            out_path = out_path / f"{name}.xlsx"

        out_path.parent.mkdir(exist_ok=True, parents=True)
        return out_path

    def create_project_manager_sheets(self,
                                      out_path: [str, Path],
                                      new_folder: str = None,
                                      sheet_name: str = "Dashboard", ):

        for pm in self.consolidator.project_managers:
            pm_df = self.consolidator.filter_by_project_manager(pm)

            # Index's will be the same as in the master, so need to reset it so it starts at row 1 again
            pm_df = reset_index(pm_df)
            pm_out_path = self.get_pm_out_path(out_path, name=pm, new_folder=new_folder)

            self._create_dashboard(pm_out_path, pm_df, sheet_name, )


if __name__ == "__main__":

    #     #TODO: Sydney trains here
    base_path = Path().cwd()
    parent_path = base_path.parent

    config_path = base_path / "config_excel.yaml"
    exclusions_path = base_path / "exclusions.yml"

    test_out = parent_path / "test"
    test_out.mkdir(exist_ok=True)

    bst_path = parent_path / "test_data" / "Project Detail.xlsx"
    prev_dash = parent_path / "test_data" / "Dashboard.xlsx"
    pm_sheets_path = parent_path / "test_data" / "Project Manager Sheets"

    dashboard = DashboardGenerator(
        config_path=config_path,
        exclusions_file_path=exclusions_path,
        issue=6
    )

    dashboard.load_bst(bst_path)

    dashboard.load_prev(prev_dash)

    dashboard.load_pm_sheets(pm_sheets_path)

    dashboard.create_consolidator()

    dashboard.update_master_with_pm()

    dashboard.create_master_dashboard(test_out / "test.xlsx")

    # dashboard.create_project_manager_sheets(test_out, new_folder="Project Manager Sheets" )

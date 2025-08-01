from modules import Agent, process_workbooks
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


if __name__ == "__main__":
    process_workbooks(Agent.Permissionaria)
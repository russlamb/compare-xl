from compare_xl import CompareXl
from sql_to_xl import SqlToXl


class CompareSqlInXl():
    """Class that connects to two databases, runs sql on each, and saves results to files.
    Then, compares the results and saves the differences to a file.  Connection strings
    passed to the class on instantiation."""
    def __init__(self, conn_str1, conn_str2):
        self.conn1=conn_str1
        self.conn2=conn_str2

    def run(self, file1, file2, diff_file, sheet_name, sql1, sql2=None):
        """Run SQL on two databases, save to Excel files, compare those files and save differences"""
        sql2 = sql1 if sql2 is None else sql2
        sql_run1 = SqlToXl(self.conn1)
        sql_run1.save_sql(sql1, file1, sheet_name)

        sql_run2 = SqlToXl(self.conn2)
        sql_run2.save_sql(sql2, file2, sheet_name)

        compare_xl = CompareXl(file1, file2, sheet_name, sheet_name)
        compare_xl.compare_and_save(diff_file)

    def run_index(self, file1, file2, diff_file, sheet_name, index, sql1, sql2=None):
        """Run SQL on two databases, save to Excel files, compare those files using an index and save differences"""
        sql2 = sql1 if sql2 is None else sql2
        sql_run1 = SqlToXl(self.conn1)
        sql_run1.save_sql(sql1, file1, sheet_name)

        sql_run2 = SqlToXl(self.conn2)
        sql_run2.save_sql(sql2, file2, sheet_name)

        compare_xl = CompareXl(file1, file2, sheet_name, sheet_name)
        compare_xl.compare_index_and_save(index,diff_file)

    @staticmethod
    def compare_only(file1, file2, diff_file, sheet_name, index=None):
        """Static Method.  No SQL is run, just compare two files and save differences.  Index optional."""
        compare = CompareXl(file1,file2,sheet_name,sheet_name)

        if index is None:
            diffs = compare.compare()
        else:
            diffs = compare.compare_index(index)

        compare.save_differences(diff_file,diffs)

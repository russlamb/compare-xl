import argparse, datetime
import configparser
from compare_sql_xl import CompareSqlInXl

if __name__ == "__main__":
    #read defaults from config.ini
    config_parser = configparser.ConfigParser()
    config_parser.read("config.ini")
    config = config_parser["DEFAULT"]

    parser = argparse.ArgumentParser(description="Run SQL against two databases and export to two excel spreadsheets." +
                                                 "  Compare two Excel files and create " +
                                                 "a third spreadsheet with differences.",
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    group = parser.add_mutually_exclusive_group(required=True)

    group.add_argument("-compare",  action="store_true", default=False,
                        help="don't run SQL, just compare the two excel files specified in -left and -right, output" +
                             " differences to -diff.  Mutually exclusive with -sql.")
    group.add_argument("-sql", #metavar="--sql_query",
                        help="SQL Query to run.  Mutually exclusive with -compare.")
    parser.add_argument("-left", #metavar="--left_file",
                        default=config["left"],
                        help="Path to workbook that will be compared.  Column in Diff file taken from this file.")
    parser.add_argument("-right", #metavar="--right_file",
                        default=config["right"],
                        help="Path to Second workbook for the comparison.")
    parser.add_argument("-diff", #metavar="--diff_file",
                        default=config["diff"],
                        help="Path to Workbook where differences will be stored.")
    parser.add_argument("-sheet", #metavar="--sheet",
                        default=config["sheet"],
                        help="name of sheet in Excel workbook where left/right data will be saved")
    parser.add_argument("-db_left", #metavar="--prod_connection",
                        default=config["db_left"],
                        help="Connection string for left side file (defaults to Production)")
    parser.add_argument("-db_right", #metavar="--qa_connection",
                        default=config["db_right"],
                        help="Database Connection string for right side file (defaults to QA database)")
    parser.add_argument("-index", #metavar="--index_column",
                        type=int,
                        help="number indicating which column to use as an index when comparing each file. " +
                             "1 specifies first column, 2 second, etc.  " +
                             "If omitted, each row is compared with the ordinal row in other file (e.g. row 1 to row " +
                             "1).  If included, rows are lined up according to index and only matching values are " +
                             "compared.  E.g. similar to an inner join.")

    parser.add_argument("--version", action="version", version='%(prog)s {}'.format(config["version"]))

    args = parser.parse_args()
    print("CLI arguments {}".format(args))

    compare = CompareSqlInXl(args.db_left, args.db_right)

    if args.compare is False and args.sql is None:
        raise ValueError("You must either pass the -compare or -sql argument.  See --help for more.")

    if args.compare is True :
        print("comparison only {} vs {}".format(args.left, args.right))
        compare.compare_only(args.left, args.right, args.diff, args.sheet, args.index)
    elif args.index is not None:
        print("running with index {}".format( datetime.datetime.now()))
        compare.run_index(args.left, args.right, args.diff, args.sheet, args.index, args.sql)

    else:
        print("running direct comparison (no index) {}".format(datetime.datetime.now()))
        compare.run(args.left, args.right, args.diff, args.sheet, args.sql)
    print("finished {}".format( datetime.datetime.now()))

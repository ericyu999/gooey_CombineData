from argparse import ArgumentParser
import os
from gooey import Gooey, GooeyParser

@Gooey(program_name='Create Quarterly Marketing Report')
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Create Quarterly Marketing Report')
    parser.add_argument('data_directory',
                        action='store',
                        default=stored_args.get('data_directory'),
                        widget='DirChooser',
                        help="Source directory that contains Excel files")
    parser.add_argument('output_directory',
                        action='store',
                        default=stored_args.get('output_directory'),
                        widget='DirChooser',
                        help="Output directory to save summary report")
    parser.add_argument('cust_file',
                        action='store',
                        default=stored_args.get('cust_file'),
                        widget='FileChooser',
                        help='Customer Account Status File')
    parser.add_argument('-d', help='Start date to include',
                        default=stored_args.get('d'),
                        widget='DateChooser'
                        )
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args


if __name__ == '__main__':
    conf = parse_args()
    print("Reading sales files")
    sales_df = combine_files(conf.data_directory)
    print("Reading customer data and combining with sales")
    customer_status_sales = add_customer_status(sales_df, conf.cust_file)
    print("Saving sales and customer summary data")
    save_results(customer_status_sales, conf.output_directory)
    print("Done")

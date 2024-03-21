import pandas as pd
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression


def read_excel_data(excel_path, sheet_name):
    """
    Read data from an Excel file and handle missing values.

    Args:
        excel_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet containing the data.

    Returns:
        pd.DataFrame: DataFrame containing the data from the Excel sheet.
                      Returns None if file not found or an error occurs.
    """
    try:
        # Read Excel data into a DataFrame and drop rows with missing values
        dataframe = pd.read_excel(excel_path, sheet_name=sheet_name)
        return dataframe.dropna()
    except FileNotFoundError:
        print("Error: File not found.")
        return None
    except Exception as e:
        print("Error:", e)
        return None


def plot_data(dataframe, x_col, y_col, title):
    """
    Plot the data and display linear regression.

    Args:
        dataframe (pd.DataFrame): DataFrame containing the data.
        x_col (str): Name of the column for the X-axis.
        y_col (str): Name of the column for the Y-axis.
        title (str): Title for the plot.
    """
    # Plot data points
    plt.plot(dataframe[x_col], dataframe[y_col], "g--", lw=3)
    plt.xlabel(x_col)
    plt.ylabel(y_col)
    plt.title(title)
    plt.grid(True)


def perform_linear_regression(dataframe, x_col, y_col):
    """
    Perform linear regression and return slope and intercept.

    Args:
        dataframe (pd.DataFrame): DataFrame containing the data.
        x_col (str): Name of the column for the X-axis.
        y_col (str): Name of the column for the Y-axis.

    Returns:
        tuple: Tuple containing slope and intercept of the regression line.
    """
    # Initialize linear regression model
    model = LinearRegression()
    # Fit the model to the data
    model.fit(dataframe[[x_col]], dataframe[[y_col]])
    # Retrieve slope and intercept of the regression line
    return model.coef_[0][0], model.intercept_[0]


def main():
    """Main function to execute the script."""
    # Prompt user for Excel file path and sheet name
    excel_path = input("Enter the Excel file path: ").strip(
        '"')  # Remove quotes from the path
    sheet_name = input("Enter the sheet name: ")

    # Read data from Excel
    data = read_excel_data(excel_path, sheet_name)
    if data is None:
        return

    # Display the loaded data
    print("Excel Data:")
    print(data)

    # Prompt user for column names and plot title
    x_column = input("Enter the column name for the X-axis: ")
    y_column = input("Enter the column name for the Y-axis: ")
    plot_title = input("Enter the title for the plot: ")

    # Plot the data
    plot_data(data, x_column, y_column, plot_title)

    # Perform linear regression and retrieve slope and intercept
    slope, intercept = perform_linear_regression(data, x_column, y_column)
    # Generate equation of the regression line
    equation = f'y = {slope:.2f}x + {intercept:.2f}'
    # Add equation to the plot
    plt.text(0.1, 0.9, equation, transform=plt.gca().transAxes,
             fontsize=12, verticalalignment='top')

    # Display the plot
    plt.show()


if __name__ == "__main__":
    main()

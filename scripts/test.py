import datetime

# Define the starting date (January 2019)
start_year = 2019
start_month = 1

# Get the current date
current_date = datetime.datetime.now()
current_year = current_date.year
current_month = current_date.month

# Calculate the total number of months
months = (current_year - start_year) * 12 + (current_month - start_month)

if __name__ == "__main__":
    print(f"Number of months from January 2019 to now: {months}")
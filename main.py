def main():
    n = 1000
    print(run_experiment_repeatedly(n))  # Simulation answer to Q1 (a).
    print(run_experiment_repeatedly(n, two_dice=True))  # Simulation answer to Q1 (b).


def roll_repeatedly_one_dice():
    """
    Roll one twelve-sided die repeatedly until all the possible
    outcomes have occurred at least once. Returns the
    number of rolls.
    """

    import random

    # boolean list determining whether each outcome has occurred.
    outcomes_occurred = [False for i in range(12)]

    result = 0
    desired_outcome = [True for i in range(12)]
    exit_condition = False

    while not exit_condition:
        roll = random.randint(1, 12)
        outcomes_occurred[roll - 1] = True
        exit_condition = outcomes_occurred == desired_outcome
        result += 1

    return result


def roll_repeatedly_two_dice():
    """
    Roll two six-sided dice repeatedly until all the possible
    sums of both dice have occurred at least once. Returns the
    number of rolls.
    """

    import random

    outcomes_occurred = [False for i in range(11)]
    desired_outcome = [True for i in range(11)]
    exit_condition = False
    result = 0

    while not exit_condition:
        roll_one = random.randint(1, 6)
        roll_two = random.randint(1, 6)
        sum_of_rolls = roll_one + roll_two

        outcomes_occurred[sum_of_rolls - 2] = True
        exit_condition = outcomes_occurred == desired_outcome
        result += 1

    return result


def run_experiment_repeatedly(n: int, two_dice=False):
    """
    Calls one of the roll_repeatedly() functions (depending on the value of the parameter
     "two_dice") n times and returns the average number of rolls.

    This function additionally plots the results of the roll_repeatedly calls along with the
    average number of rolls, and writes the data into an Excel sheet.
    """

    assert(n >= 0)

    import matplotlib.pyplot as plt
    from openpyxl import Workbook
    from openpyxl.styles import Font

    workbook = Workbook()
    sheet = workbook.active

    experiment_results = []

    sheet["A1"] = "Experiment Number"
    sheet["A1"].font = Font(bold=True)
    sheet["B1"] = "No. of rolls"
    sheet["B1"].font = Font(bold=True)

    for i in range(n):
        experiment_outcome = 0
        if two_dice:
            experiment_outcome = roll_repeatedly_two_dice()
        else:
            experiment_outcome = roll_repeatedly_one_dice()
        experiment_results.append(experiment_outcome)

        sheet[f"A{i + 2}"] = i + 1
        sheet[f"B{i + 2}"] = experiment_outcome

    average = sum(experiment_results) / len(experiment_results)

    sheet["C1"] = "Avg no. of rolls"
    sheet["C1"].font = Font(bold=True)
    sheet["C2"] = average

    plt.plot(experiment_results)
    plt.axhline(average, color="red")
    if two_dice:
        plt.title(f"Running two-dice experiment {n} times")
    else:
        plt.title(f"Running one-dice experiment {n} times")
    plt.show()

    plt.hist(experiment_results)

    if two_dice:
        workbook.save(filename="rolls_two_dice.xlsx")
        plt.title("Histogram plot of two-dice experiment")
    else:
        workbook.save(filename="rolls_one_dice.xlsx")
        plt.title("Histogram plot of one-dice experiment")
    plt.show()

    return average


if __name__ == "__main__":
    main()
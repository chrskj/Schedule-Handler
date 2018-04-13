from __future__ import print_function
import sys
from ortools.constraint_solver import pywrapcp
import os
import pandas as pd


def get_shift_array(schedule):

    shift_array_array = []
    people_array = []

    for i in range(6, schedule.shape[0] - 1):

        shift_array = []
        person_schedule = schedule.iloc[i,:]

        for j in range(schedule.shape[1]):
            if isinstance(person_schedule[j], float):
                shift_array.append(0)
            elif person_schedule[j] == "(OK)":
                shift_array.append(1)
            elif person_schedule[j] == "OK":
                shift_array.append(2)
            else:
                people_array.append(person_schedule[j])

        shift_array_array.append(shift_array)

    return [people_array, shift_array_array]


def get_num_on_shift(schedule):

    num_on_shift = []
    person_schedule = schedule.iloc[2,:]

    for j in range(1, schedule.shape[1]):
        num_on_shift.append(person_schedule[j])

    return num_on_shift


def get_date_format(schedule):

    date_format = []
    dager = schedule.iloc[4, :]
    tider = schedule.iloc[5, :]

    for j in range(1, schedule.shape[1]):
        if isinstance(dager[j], float):
            date_format.append(dager[j - 1] + ": " + tider[j])
        else:
            date_format.append(dager[j] + ": " + tider[j])

    return date_format


def get_week_data(schedule):

    week_start = [0]
    week_end = []
    end = 0

    uker = schedule.iloc[1, :]

    for j in range(2, schedule.shape[1]):
        if isinstance(uker[j], int):
            week_start.append(j - 1)
            week_end.append(j - 2)
        end = j - 1

    week_end.append(end)

    return [week_start, week_end]


def main():
    # File setup
    print("Automatisk generering av vaktlister")
    file = 'Doodle.xls'
    doodle = pd.read_excel(file, header=None)

    # Creates the solver.
    solver = pywrapcp.Solver("simple_example")

    # Cook availability
    doodle_form = get_shift_array(doodle)
    cook_names = doodle_form[0]
    cook_avail = doodle_form[1]

    # Variables
    num_cooks = len(cook_names)
    num_shifts = len(cook_avail[0])
    num_on_shifts = get_num_on_shift(doodle)
    num_weeks = len(get_week_data(doodle)[0])
    week_start = get_week_data(doodle)[0]
    week_end = get_week_data(doodle)[1]
    date_format = get_date_format(doodle)

    cook_working = [[] for i in range(num_weeks)]
    cook_not_working = [[] for i in range(num_weeks)]

    # shift variable
    shifts = {}
    for i in range(num_shifts):
        for j in range(num_cooks):
            shifts[(i, j)] = solver.IntVar(0, 1, "shifts(%i, %i)" % (i, j))
    # shifts[(i, j)] is the status of shift i for cook j
    shifts_flat = [shifts[(i, j)] for i in range(num_shifts) for j in range(num_cooks)]

    # Each cook works one time or less each week
    for j in range(num_cooks):
        for week in range(num_weeks):
            solver.Add(solver.Sum([shifts[(i, j)] for i in range(week_start[week], week_end[week] + 1)]) <= 1)

    # Each shift has a certain amount of cooks working
    for i in range(num_shifts):
        num_on_shift = num_on_shifts[i]
        solver.Add(solver.Sum([shifts[(i, j)] for j in range(num_cooks)]) == num_on_shift)

    # Each cook can only work certain days
    for j in range(num_cooks):
        for i in range(num_shifts):
            if cook_avail[j][i] == 0:
                solver.Add(shifts[(i, j)] == 0)

    # Create the decision builder.
    db = solver.Phase(shifts_flat, solver.CHOOSE_FIRST_UNBOUND,
                      solver.ASSIGN_MIN_VALUE)

    # Create the solution collector.
    solution = solver.Assignment()
    solution.Add(shifts_flat)
    collector = solver.AllSolutionCollector(solution)
    solutions_limit = solver.SolutionsLimit(1)
    solver.Solve(db, [solutions_limit, collector])

    # Display a few solutions picked at random.
    print("Solutions found:", collector.SolutionCount())
    print("Time:", solver.WallTime(), "ms")
    a_few_solutions = [0]
    with open('out.txt', 'w') as f:
        print("Solutions found:", collector.SolutionCount(), file=f)
        print("Time:", solver.WallTime(), "ms", file=f)
        for sol in a_few_solutions:
            print("Solution number", sol, file=f)

            for i in range(num_shifts):
                print("\n", file=f)
                print(date_format[i], file=f)
                for j in range(num_cooks):
                    result = collector.Value(sol, shifts[(i, j)])
                    if result == 1:
                        print(cook_names[j], file=f)

            for week in range(num_weeks):
                for i in range(week_start[week], week_end[week] + 1):
                    for j in range(num_cooks):
                        result = collector.Value(sol, shifts[(i, j)])
                        if result == 1:
                            cook_working[week].append(cook_names[j])

                for j in range(num_cooks):
                    if cook_names[j] not in cook_working[week]:
                        cook_not_working[week].append(cook_names[j])

        print("\n", file=f)
        print("Folk som ikke er satt opp:", file=f)
        for week in range(num_weeks):
            print("\n", file=f)
            print("Uke", week, file=f)
            for i in range(len(cook_not_working[week])):
                #print("\n", file=f)
                print(cook_not_working[week][i], file=f)


if __name__ == "__main__":
    main()

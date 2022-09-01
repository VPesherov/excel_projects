
def to_fixed(num_obj, digits=0):
    return f"{num_obj:.{digits}f}"


def city_from_brackets(my_string):
    i = 0
    result_string = ""
    while i < len(my_string):
        if my_string[i] == '(':
            result_string = my_string[i + 1: len(my_string) - 1:]
            break
        i += 1
    return result_string

def check_dissmised(orange_dict):
    y= 0
    n= 0

    not_elloyy_dublicate = {}

    for key, value in orange_dict.items():
        
        date: str = value[0]['Увольн.']
        if date.startswith("2025") or date.startswith("2026") :
            y+=1
            # print(key)
            # print(date)
            not_elloyy_dublicate[key] = value
        else:
            n+=1

    # print("уволеных в 2025", y)
    # print("уволеных в другие года ", n)
    return not_elloyy_dublicate
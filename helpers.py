def letter(num: int):
    letters =['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    length = len(letters)
    rest_ind = num % length - 1

    if num<length: return letters[rest_ind]
    else:
        return letters[num//length-1]+letters[rest_ind]
     
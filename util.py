def split_array(arr: list, chunk_size: int):
    return [arr[i:i + chunk_size] for i in range(0, len(arr), chunk_size)]
def str_select(str1: str, str2: str):
    return str1 if str1 != '' else str2

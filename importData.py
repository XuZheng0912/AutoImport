def import_check_data(username: str, password: str, mode: int):
    switcher = {
        0: import_new,
        1: import_cover
    }
    importer = switcher.get(mode)
    importer(username, password)


def import_new(username: str, password: str):
    print(username, password)


def import_cover(username: str, password: str):
    print(username, password)



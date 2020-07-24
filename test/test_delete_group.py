import random


def test_delete_group(app):
    if len(app.groups.get_group_list()) < 2:
        app.groups.add_new_groups("New Test Group")
    old_list = app.groups.get_group_list()
    index_group = random.choice(old_list)
    index = old_list.index(index_group)
    print(index)
    app.groups.delete_groups(index)
    new_list = app.groups.get_group_list()
    old_list.remove(index_group)
    assert sorted(old_list) == sorted(new_list)
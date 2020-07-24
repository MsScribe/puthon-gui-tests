def test_add_group(app):
    old_list = app.groups.get_group_list()
    app.groups.add_new_groups("Name group")
    new_list = app.groups.get_group_list()
    old_list.append("Name group")
    assert sorted(old_list) == sorted(new_list)
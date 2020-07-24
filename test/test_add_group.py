def test_add_group(app, data_groups):
    group = data_groups
    print(group)
    old_list = app.groups.get_group_list()
    app.groups.add_new_groups(group)
    new_list = app.groups.get_group_list()
    old_list.append(group)
    assert sorted(old_list) == sorted(new_list)
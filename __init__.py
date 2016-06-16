from contextlib import contextmanager
try:
    import _winreg as wr
except Exception, e:
    LOG.error("Can't load win reg to register commands: %s" % e)

@contextmanager
def closing(root, path, style, value=None):
    open_key = None
    try:
        open_key = wr.OpenKey(root, path, 0, style)
        yield open_key
    except:
        reg = wr.CreateKey(root, path)
        if value:
            wr.SetValueEx(reg, None, 0, wr.REG_SZ, value)
        wr.CloseKey(reg)
        open_key = wr.OpenKey(root, path, 0, style)
        yield open_key
    finally:
        if open_key is not None:
            wr.CloseKey(open_key)
 
def register_extension(ext=None, name="Open with python", command="python %1", shift_required=False):
    """
    register_extension allows you to have python do registry editing to add commands to the context menus in Windows Explorer.
    
    :param str|None ext: String extension that you want to be affected by this change. None if you want it to affect all.
    :param str name: User facing display name for this.
    :param str command: The command that you want to run when the menu item is clicked. %1 is the object that you have selected being passed in.
    :param bool shift_required: Does this command require shift to be held to show up?
    :return: True if it succeeded.
    :rtype: bool
    """
    regPath = ext
    if ext is None:
        regPath = "*"
   
    with closing(wr.HKEY_CLASSES_ROOT, regPath, wr.KEY_READ, value=ext[1].upper() + ext[2:] + "_pyReg") as key:
        val = wr.QueryValue(key, None)
   
    with closing(wr.HKEY_CLASSES_ROOT, val, wr.KEY_READ) as key:
        with closing(key, "shell", wr.KEY_READ) as key_shell:
            with closing(key_shell, name, wr.KEY_READ) as key_title:
                if shift_required:
                    wr.SetValueEx(reg, "Extended", 0, wr.REG_SZ)
                with closing(key_title, "command", wr.KEY_SET_VALUE) as key_command:
                    wr.SetValueEx(key_command, None, 0, wr.REG_SZ, command)
    return True

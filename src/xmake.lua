target("portable-file-dialogs")
    add_rules("module.component")
    add_files("portable-file-dialogs/portable-file-dialogs.cpp")

    add_defines("PFD_SKIP_IMPLEMENTATION=1", {interface = true})
    add_includedirs(path.join(os.scriptdir(), "portable-file-dialogs"), {interface = true})

    if is_plat("windows") then
        add_syslinks("user32", "shell32", "advapi32")
    end

target("star-tool")
    add_rules("module.program")
    add_files("*.cpp")

    if is_plat("windows") then
        add_defines("WIN32_LEAN_AND_MEAN")
    end

    add_deps("portable-file-dialogs")

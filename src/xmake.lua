target("star-tool")
    add_rules("module.program")
    add_files("*.cpp")

    if is_plat("windows") then
        add_defines("WIN32_LEAN_AND_MEAN")
    end

    add_packages("nativefiledialog-extended")

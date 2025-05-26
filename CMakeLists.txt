cmake_minimum_required(VERSION 3.27)
project(sendto_recomposed C)
add_definitions(-D_WIN32_WINNT=0x0601)

set(CMAKE_C_STANDARD 11)

# Enable resource file compilation
enable_language(RC)

# Use Windows subsystem (no console window)
if(MSVC)
    set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} /SUBSYSTEM:WINDOWS")
endif()

# Source and resource files
set(SRC      ${CMAKE_CURRENT_SOURCE_DIR}/sendto.c)
set(RC_FILE  ${CMAKE_CURRENT_SOURCE_DIR}/sendto.rc)
set(MANIFEST ${CMAKE_CURRENT_SOURCE_DIR}/sendto.manifest)

# Define executable target and output name as sendto.exe
add_executable(sendto_recomposed WIN32 ${SRC} ${RC_FILE})
set_target_properties(sendto_recomposed PROPERTIES OUTPUT_NAME "sendto")

# Unicode support
target_compile_definitions(sendto_recomposed PRIVATE UNICODE _UNICODE)

# MSVC compile options per build type
if(MSVC)
    target_compile_options(sendto_recomposed PRIVATE
            /EHsc
            # Debug settings
            $<$<CONFIG:Debug>:/MDd /Zi /RTC1>
            # Release settings
            $<$<CONFIG:Release>:/MD /O2>
    )
endif()

# Link required Windows libraries
target_link_libraries(sendto_recomposed PRIVATE
        ole32
        shell32
        shlwapi
        comctl32
        user32
        gdi32
        uuid
)

# Embed manifest into the generated sendto.exe
add_custom_command(TARGET sendto_recomposed POST_BUILD
        COMMAND mt.exe -nologo -manifest ${MANIFEST} -outputresource:$<TARGET_FILE:sendto_recomposed>;#1
        COMMENT "Embedding manifest into sendto.exe"
)

# Create a 'sendto' directory alongside sendto.exe
add_custom_command(TARGET sendto_recomposed POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:sendto_recomposed>/sendto"
        COMMENT "Creating 'sendto' directory next to sendto.exe"
)

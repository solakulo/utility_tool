#include <stdio.h>
#include <time.h>
#include <direct.h> // For _mkdir on Windows

int main() {
    // Get the current date
    time_t t = time(NULL);
    struct tm* tm_info = localtime(&t);
    char folder_name[20];
    strftime(folder_name, sizeof(folder_name), "%Y-%m-%d", tm_info);

    // Create the folder
    if (_mkdir(folder_name) == 0) {
        printf("Folder '%s' created successfully!\\n", folder_name);
    }
    else {
        printf("Failed to create folder '%s'. It may already exist.\\n", folder_name);
    }

    return 0;
}
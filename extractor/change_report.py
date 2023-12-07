import os


class ChangeReport:

    def __init__(self, base_folder_path):
        self.base_folder_path = base_folder_path
        self.newest_folder, self.previous_folder = self.get_recent_folders()
        self.newest_folder_contents = ChangeReport.get_folder_contents(self.newest_folder_path)

        if self.previous_folder:
            self.previous_folder_contents = ChangeReport.get_folder_contents(self.previous_folder_path)
            self.added_set, self.deleted_set = self.get_added_deleted()
            self.changed_set = self.get_changed()
        else:
            self.previous_folder = "None"
            self.added_set = set(self.newest_folder_contents)
            self.deleted_set = set()
            self.changed_set = set()

    @property
    def newest_folder_path(self):
        return ChangeReport.get_absolute_path(self.base_folder_path, self.newest_folder)

    @property
    def previous_folder_path(self):
        return ChangeReport.get_absolute_path(self.base_folder_path, self.previous_folder)

    @property
    def shared_contents(self):
        return set(self.newest_folder_contents) - self.added_set - self.deleted_set

    def get_recent_folders(self):
        # Get a list of all directories in the specified path
        dirs = [d for d in os.listdir(self.base_folder_path) if os.path.isdir(os.path.join(self.base_folder_path, d))]
        dirs += ["", ""]

        # Sort the directories by creation time, most recent first
        dirs.sort(key=lambda x: os.path.getctime(os.path.join(self.base_folder_path, x)), reverse=True)

        # Return the two most recent folder names (or "" if NA)
        return dirs[0], dirs[1]

    def get_added_deleted(self):
        newest_set = set(self.newest_folder_contents)
        previous_set = set(self.previous_folder_contents)

        return newest_set - previous_set, previous_set - newest_set

    def get_changed(self):
        updated = set()

        for relative_path in self.shared_contents:
            with open(ChangeReport.get_absolute_path(self.newest_folder_path, relative_path)) as newest_file, \
                    open(ChangeReport.get_absolute_path(self.previous_folder_path, relative_path)) as previous_file:
                if newest_file.read() != previous_file.read():
                    updated.add(relative_path)

        return updated

    def write_change_report(self):
        change_report_title = f"Change Report{' (DELETIONS DETECTED)' if self.deleted_set else ''}.txt"
        deleted_lines = [f"\t{line}" for line in self.deleted_set] if self.deleted_set else ["\tNone"]
        added_lines = [f"\t{line}" for line in self.added_set] if self.added_set else ["\tNone"]
        changed_lines = [f"\t{line}" for line in self.changed_set] if self.changed_set else ["\tNone"]

        with open(ChangeReport.get_absolute_path(
            f"{self.base_folder_path}/{self.newest_folder}", change_report_title
        ), "w") as change_report:
            results = [
                "CHANGE REPORT",
                f"{self.previous_folder} --> {self.newest_folder}",
                "",
                "DELETED FILES:",
                *deleted_lines,
                "",
                "ADDED FILES:",
                *added_lines,
                "",
                "CHANGED FILES:",
                *changed_lines,
            ]
            change_report.write('\n'.join(results))

    @staticmethod
    def get_absolute_path(base_path, relative_path):
        return f"{base_path}/{relative_path}"

    @staticmethod
    def get_folder_contents(folder_path, original_path=None):
        folder_contents = []
        if original_path is None:
            original_path = f"{folder_path}/"

        for element in os.listdir(folder_path):
            if os.path.isdir(os.path.join(folder_path, element)):
                folder_contents.extend(
                    ChangeReport.get_folder_contents(
                        ChangeReport.get_absolute_path(folder_path, element),
                        original_path,
                    )
                )
            elif "Change Report" not in element:
                folder_contents.append(
                    ChangeReport.get_absolute_path(folder_path, element).replace(original_path, "")
                )

        return folder_contents

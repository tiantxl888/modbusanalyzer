import json
import os

class ProjectManager:
    def __init__(self, project_file='project.json'):
        self.project_file = project_file

    def save(self, config: dict):
        with open(self.project_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def load(self) -> dict:
        if not os.path.exists(self.project_file):
            return {}
        with open(self.project_file, 'r', encoding='utf-8') as f:
            return json.load(f)

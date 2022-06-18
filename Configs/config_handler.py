import yaml

class config_handler:
    
    def __init__(self, fileName):
        with open("Configs/" + fileName, 'r', encoding="utf8") as f:
            self.config_data = yaml.safe_load(f)
    
    def get_user_name(self):
        try:
            return self.config_data['user_name']
        except yaml.YAMLError as exc:
            # Write to logs
            print(exc)
            return None
    
    def get_folder_name(self):
        try:
            return self.config_data["Folder_name"]
        except yaml.YAMLError as exc:
            # Write to logs
            print(exc)
            return None
    
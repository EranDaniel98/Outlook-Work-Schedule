import yaml

class config_handler:
    
    def __init__(self, fileName):
        try:
            with open("Configs/" + fileName, 'r', encoding="utf8") as f:
                self.config_data = yaml.safe_load(f)
        except yaml.YAMLError as exc:
            # Write to log
            print(exc)
            return None
    
    def get_requested_param(self, param):
        try:
            return self.config_data[param]
        except yaml.YAMLError as exc:
            # Write to logs
            print(exc)
            return None
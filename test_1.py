import re
timestamp = re.match(r'^\d+', "20241210111215023044_Naukri_AmitSinghal[13y_0m].pdf").group()
print(type(timestamp),timestamp)
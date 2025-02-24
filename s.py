if page_data["Compte à débiter"] == None or page_data["Compte à débiter"] == "":
            alternative_match = re.findall(r"Mail[\s\S]*?\b(\d{2}", page_content)
            if alternative_match:
                page_data["Compte à débiter"] = alternative_match[0]

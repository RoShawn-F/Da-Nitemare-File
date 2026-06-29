### Da-Nitemare-File
### An optional challenge that created by the instructors at VetsInTech...
### it was cool to do but it was exhausting (or Maybe I just need to go to bed) ^_^

next_url = character_url
all_characters = []

while next_url:
    next_response = requests.get(next_url)
    next_data = next_response.json()
    all_characters.extend(next_data["results"])
    next_url = next_data["info"]["next"]
    print(len(all_characters))
    time.sleep(0.5)


### set up a workbook and worksheet titled "Rick and Morty Characters"

print(data["results"][0].keys())

headers = ["id", "name", "status", "species", "type", "gender", "origin", "location", "image", "episode", "url", "created"]

### assign a variable 'data' with the returned GET request
for col_num, header in enumerate(headers, 1):
    ws.cell(row=1, column=col_num, value=header)
### create the appropriate headers in openpyxl for all of the keys for a single character
for row_num, character in enumerate(all_characters, 2):
    if character["episode"]: 
        print("processing episodes for:", character["name"])
        current_episode = []
        for episode_url in character["episode"]:
            try:
                ep_response = requests.get(episode_url)
                ep_data = ep_response.json()
                current_episode.append(ep_data["name"])
                time.sleep(0.3)
            except:
                current_episode.append(episode_url)
        character["episode"] = current_episode
    for col_num, value in enumerate(character.values(), 1):
        ws.cell(row=row_num, column=col_num, value=str(value))

### loop through all of the 'results' of the data to populate the rows and columns for each character

### NOTE: due to the headers, the rows need to be offset by one!

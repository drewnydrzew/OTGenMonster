import pandas as pd
import random
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom.minidom import parseString
import json

file_path = 'input.ods'
output_file_path = 'output.ods'
data = pd.read_excel(file_path, engine='odf')

monster_attacks_output = []
monster_defenses_output = []

all_monsters_df = pd.DataFrame()


min_values = data.iloc[0].to_dict()
max_values = data.iloc[1].to_dict()

def create_monster_df_row(name, hp, exp, race, speed, dps, attacks_count, attacks,
                          armor_defense, defenses_count, defense, loot, monster_xml_tfs, monster_xml_rme):
    """
    Create a DataFrame row for a new monster.
    
    :param name: Monster's name
    :param hp: Monster's health points
    :param exp: Experience points given by the monster
    :param race: Monster's race
    :param speed: Monster's speed
    :param dps: Damage per second of the monster
    :param attacks_count: Number of attacks
    :param attacks: Descriptions of attacks
    :param armor_defense: Combined value of armor and defense
    :param defenses_count: Number of defenses
    :param defense: Descriptions of defenses
    :param loot: Loot dropped by the monster
    :return: DataFrame row with the monster's data
    """
    data = {
        'NAME': [name],
        'HP': [hp],
        'EXP': [exp],
        'RACE': [race],
        'SPEED': [speed],
        'DPS': [dps],
        'NUMBER OF ATTACKS': [attacks_count],
        'ATTACKS': [attacks],
        'ARMOR + DEFENSE': [armor_defense],
        'NUMBER OF DEFENSES': [defenses_count],
        'DEFENSE': [defense],
        'LOOT': [loot],
        'MONSTER_XML_TFS': [monster_xml_tfs],
        'MONSTER_XML_RME': [monster_xml_rme]
    }
    return pd.DataFrame(data)



def generate_loot(base_loot_json_min, base_loot_json_max, additional_loot_json):
    base_loot_min = base_loot_json_min if isinstance(base_loot_json_min, dict) else json.loads(base_loot_json_min or '[]')
    base_loot_max = base_loot_json_max if isinstance(base_loot_json_max, dict) else json.loads(base_loot_json_max or '[]')
    additional_loot = additional_loot_json if isinstance(additional_loot_json, dict) else json.loads(additional_loot_json or '[]')
    
    # Generate the base loot by randomizing between min and max values
    base_loot = []
    for min_item, max_item in zip(base_loot_min, base_loot_max):
        if "name" in min_item:
            item = {
                "name": min_item["name"],
                "countmax": random.randint(min_item["countmax"], max_item["countmax"]),
                "chance": random.randint(min_item["chance"], max_item["chance"])
            }
        else:
            item = {
                "id": min_item["id"],
                "countmax": random.randint(min_item["countmax"], max_item["countmax"]),
                "chance": random.randint(min_item["chance"], max_item["chance"])
            }
        base_loot.append(item)
    
    # Combine base loot and additional loot
    combined_loot = base_loot + additional_loot
    
    return combined_loot

def generate_elements(monster, elements_min, elements_max):
    elements_count = random.randint(0, 6)  
    
    if elements_count != 0:
        elements_element = SubElement(monster, 'elements')
        for _ in range(elements_count):
                element_type = random.choice(["holy", "death", "fire", "ice", "energy", "earth"])  
                element_percent = random.randint(elements_min, elements_max)  
                SubElement(elements_element, 'element', {f'{element_type}Percent': str(element_percent)})


        return elements_element

def add_defenses_to_monster(monster, hp, speed, armor):
    defenses_element = SubElement(monster, 'defenses', {'armor': str(armor), 'defense': str(armor)})

    # Healing
    if random.choice([True, False]):  
        max_healing = int(0.1 * hp)  
        SubElement(defenses_element, 'defense', {
            'name': "healing",
            'interval': str(random.randint(1000, 2500)),
            'chance': str(random.randint(5, 15)),
            'min': str(int(0.05 * max_healing)),  
            'max': str(max_healing),
        })
        monster_defenses_output.append(["healing","max: ", max_healing])
    # Speed
    if random.choice([True, False]):  
        speedchange = random.randint(int(speed * 1.2), 1000)
        duration = random.randint(5000, 60000)
        speed_defense = SubElement(defenses_element, 'defense', {
            'name': "speed",
            'interval': str(random.randint(1000, 2500)),
            'chance': str(random.randint(5, 15)),
            'speedchange': str(speedchange),
            'duration': str(duration),
        })
        monster_defenses_output.append(["speed", "speedchange:", speedchange, "duration: ", duration])
        SubElement(speed_defense, 'attribute', {'key': "areaEffect", 'value': "redshimmer"})

    # Invisible
    if random.choice([True, False]):  
        duration = random.randint(5000, 60000)
        invisible_defense = SubElement(defenses_element, 'defense', {
            'name': "invisible",
            'interval': str(random.randint(1000, 2500)),
            'chance': str(random.randint(5, 15)),
            'duration': str(duration),
        })
        monster_defenses_output.append(["invisible", "duration: ", duration])
        SubElement(invisible_defense, 'attribute', {'key': "areaEffect", 'value': "blueshimmer"})

    # Outfit
    if random.choice([True, False]): 
        duration = random.randint(5000, 60000)
        monster_names = data['NAME'].dropna().tolist()
        monster_name = random.choice(monster_names)
        outfit_defense = SubElement(defenses_element, 'defense', {
            'name': "outfit",
            'interval': str(random.randint(1000, 2500)),
            'chance': str(random.randint(5, 15)),
            'monster': monster_name,
            'duration': str(duration),
        })
        monster_defenses_output.append(["outfit: ", monster_name, "duration: ", duration])

       
def distribute_dps(dps, attacks_count):

    percentages = sorted([random.random() for _ in range(attacks_count - 1)] + [0, 1])
    return [round((percentages[i+1] - percentages[i]) * dps) for i in range(attacks_count)]

def add_attacks_to_monster(monster, dps, attacks_count):
    area_radius_attacks = [ "physical", "earth", "ice", "energy", "fire", "death", "holy",
                            "poison", "drown", "lifedrain", "manadrain", "speed", "drunk", "outfit",
                            "poisoncondition", "freezecondition", "firecondition", "energycondition", "drowncondition",
                            "bleedcondition", "betrayed wraith skill reducer", "drown"]


    area_wave_attacks = area_radius_attacks 

    dps_distribution = distribute_dps(dps, attacks_count)

    add_melee_attack(monster, dps_distribution[0], 1)  

    if attacks_count > 1:
        for i in range(1, attacks_count):
            attack_type = random.choice(["radius", "wave"])
            if attack_type == "radius":
                add_area_radius_attack(monster, random.choice(area_radius_attacks), dps_distribution[i])
            elif attack_type == "wave":
                add_area_wave_attack(monster, random.choice(area_wave_attacks), dps_distribution[i])


def add_melee_attack(monster, dps, attacks_count):
    attack_element = SubElement(monster, 'attack', {
        'name': "melee",
        'interval': str(random.randint(500, 2500)),
        'min': str(-int(dps / 4)),
        'max': str(-dps),
          })
    monster_attacks_output.append(["melee", dps])


def add_area_radius_attack(monster, attack_name, dps):

    if attack_name == "outfit":
        monster_names = data['NAME'].dropna().tolist()

        attack_element = SubElement(monster, 'attack', {
            'name': attack_name,
            'interval': str(random.randint(500, 2500)),
            'chance': str(random.randint(5, 15)),
            'monster': random.choice(monster_names),
            'duration': str(random.randint(3000, 60000)),
            'range': str(random.randint(1, 7)),
            'radius': str(random.randint(0, 9)),
        })
    if attack_name == "drunk":
        attack_element = SubElement(monster, 'attack', {
            'name': attack_name,
            'interval': str(random.randint(500, 2500)),
            'chance': str(random.randint(5, 15)),
            'duration':  str(random.randint(3000, 60000)),
            'range': str(random.randint(1, 7)),
            'radius': str(random.randint(0, 9)),
        })

    else:
        attack_element = SubElement(monster, 'attack', {
            'name': attack_name,
            'interval': str(random.randint(500, 2500)),
            'chance': str(random.randint(5, 15)),
            'min': str(-int(dps / 4)),
            'max': str(-dps),
            'range': str(random.randint(1, 7)),
            'radius': str(random.randint(0, 9)),
        })
    monster_attacks_output.append([attack_name, dps])

def add_area_wave_attack(monster, attack_name, dps):

    if attack_name == "outfit":
        monster_names = data['NAME'].dropna().tolist()
        attack_element = SubElement(monster, 'attack', {
            'name': attack_name,
            'interval': str(random.randint(500, 2500)),
            'chance': str(random.randint(5, 15)),
            'min': str(-int(dps / 4)),
            'max': str(-dps),
            'monster': random.choice(monster_names),
            'duration': str( random.randint(3000, 60000)),
            'length': str(random.randint(1, 8)),
            'spread': str(random.randint(0, 3)),
        })
    if attack_name == "drunk":
        attack_element = SubElement(monster, 'attack', {
            'name': attack_name,
            'interval': str(random.randint(500, 2500)),
            'chance': str(random.randint(5, 15)),
            'duration':  str(random.randint(3000, 60000)),
            'length': str(random.randint(1, 8)),
            'spread': str(random.randint(0, 3)),
        })
    else:
        attack_element = SubElement(monster, 'attack', {
            'name': attack_name,
            'interval': str(random.randint(500, 2500)),
            'chance': str(random.randint(5, 15)),
            'min': str(-int(dps / 4)),
            'max': str(-dps),
            'length': str(random.randint(1, 8)),
            'spread': str(random.randint(0, 3)),
        })
    monster_attacks_output.append([attack_name, dps])



def generate_monster(row, min_values, max_values):
    name = row['NAME'] if pd.notna(row['NAME']) else "Unnamed Monster"
    hp = int(row['HP']) if pd.notna(row['HP']) else random.randint(int(min_values['HP']), int(max_values['HP']))
    exp = int(row['EXP']) if pd.notna(row['EXP']) else random.randint(int(min_values['EXP']), int(max_values['EXP']))
    attacks_count = int(row['ATTACKS']) if pd.notna(row['ATTACKS']) else random.randint(int(min_values['ATTACKS']), int(max_values['ATTACKS']))
    dps = int(row['DPS']) if pd.notna(row['DPS']) else random.randint(int(min_values['DPS']), int(max_values['DPS']))
    armor = int(row['ARMOR']) if pd.notna(row['ARMOR']) else random.randint(int(min_values['ARMOR']), int(max_values['ARMOR']))
    speed = int(row['SPEED']) if pd.notna(row['SPEED']) else random.randint(int(min_values['SPEED']), int(max_values['SPEED']))
    armor_defense = int(row['DEFENSE']) if pd.notna(row['DEFENSE']) else random.randint(int(min_values['DEFENSE']), int(max_values['DEFENSE']))
    race = random.choice(["blood", "venom", "fire", "undead", "energy"])
    base_loot_json_min = data.at[0, 'LOOT'] if pd.notna(data.at[0, 'LOOT']) else '[]'
    base_loot_json_max = data.at[1, 'LOOT'] if pd.notna(data.at[1, 'LOOT']) else '[]'
    elements_min = data.at[0, 'ELEMENTS'] if pd.notna(row['ELEMENTS']) else (min_values['ELEMENTS'])
    elements_max = data.at[1, 'ELEMENTS'] if  pd.notna(row['ELEMENTS']) else (max_values['ELEMENTS'])


    additional_loot_json = row['EXTRA LOOT'] if pd.notna(row['EXTRA LOOT']) else '[]'
    loot = generate_loot(base_loot_json_min, base_loot_json_max, additional_loot_json)
   



    monster = Element('monster', {
        'name': name,
        'race': race,
        'experience': str(exp),
        'speed': str(speed),
        'manacost': "200"  
    })
    
    health = SubElement(monster, 'health', {
        'now': str(hp),
        'max': str(hp)
    })
    look = SubElement(monster, 'look', {
        'type': "17",
        'corpse': "9962"
    })

    targetchange = SubElement(monster, 'targetchange', {
        'interval': "5000",
        'chance': "0"
    })

    flags = SubElement(monster, 'flags')
    for flag in ['summonable', 'attackable', 'hostile', 'illusionable', 'convinceable', 
                 'pushable', 'canpushitems', 'canpushcreatures', 'targetdistance', 
                 'staticattack', 'runonhealth', 'canwalkonenergy', 'canwalkonfire', 'canwalkonpoison']:
        SubElement(flags, 'flag', {flag: "1" if flag == 'attackable' or flag == 'targetdistance' or flag == 'hostile' else "0"})

    attacks_container = SubElement(monster, 'attacks')

    add_attacks_to_monster(attacks_container, dps, attacks_count)

    
    add_defenses_to_monster(monster, hp, speed, armor)

    generate_elements(monster, elements_min, elements_max)


    if pd.notna(row['SENTENCES']) and not row['SENTENCES'].startswith('MIN VALUES') and not row['SENTENCES'].startswith('MAX VALUES'):
        voices = SubElement(monster, 'voices', {
        'interval': str(5000),
        'chance': str(10)
    })
        sentences = row['SENTENCES'].split(';')
        for sentence in sentences:
            clean_sentence = sentence.strip()
            if clean_sentence:  
                SubElement(voices, 'voice', {'sentence': clean_sentence})
            



    loot_element = SubElement(monster, 'loot')
    for item in loot:
        if "name" in item:
            item_attributes = {
                "name": item["name"],
                'countmax': str(item['countmax']),
                'chance': str(item['chance'])
            }
        elif "id" in item:
            item_attributes = {
                "id": item["id"],
                'countmax': str(item['countmax']),
                'chance': str(item['chance'])
            }
        SubElement(loot_element, 'item', item_attributes)

           


    monster_xml_tfs = f'<monster name="{name}" file="folder/{name.lower().replace(" ", "_")}.xml" />'

    monster_xml_rme = f'<creature name="{name}" type="monster" looktype="17" />'

   

    monster_row = create_monster_df_row(name, hp, exp, race, speed, dps, attacks_count, monster_attacks_output, armor_defense, len(monster_defenses_output), monster_defenses_output, loot, monster_xml_tfs, monster_xml_rme)
    
    raw_xml = tostring(monster, 'utf-8', method='xml')

    pretty_xml_as_string = parseString(raw_xml).toprettyxml(indent="    ", encoding='UTF-8')
   
    return pretty_xml_as_string, name.lower().replace(" ", "_"), monster_row




for i, row in data.iterrows():
    if pd.isna(row['NAME']) or i < 2:  
        continue
    monster_xml_bytes, file_name, monster_row = generate_monster(row, min_values, max_values)
    monster_xml_str = monster_xml_bytes.decode('utf-8')  
    file_path = f'monsters/{file_name}.xml'
    with open(file_path, 'w', encoding='utf-8') as file:  # Ensure the file is opened with utf-8 encoding
        file.write(monster_xml_str)  # Write the decoded string, not bytes
    print(f'Monster "{row["NAME"]}" saved to: {file_path}')
    all_monsters_df = pd.concat([all_monsters_df, monster_row], ignore_index=True)
    monster_attacks_output = []
    monster_defenses_output = []
# Move saving the DataFrame to outside the loop to write it once after all rows are added
all_monsters_df.to_excel(output_file_path, engine='odf', index=False)

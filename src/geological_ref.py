def create_database(db_connection):
    if db_connection is None:
        raise ValueError("Database connection is not established.")

    cursor = db_connection.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS collar_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pdh_id TEXT NOT NULL,
            hole_id TEXT NOT NULL,
            location_x DECIMAL(7,2),
            location_y DECIMAL(7,2),
            location_z DECIMAL(7,2),
            azim INT,
            dip INT,
            target_length DECIMAL(7,2),
            actual_length DECIMAL(7,2),
            date_started TEXT NOT NULL,
            date_completed TEXT NOT NULL,
            sup_geo TEXT NOT NULL,
            datum TEXT,
            projection TEXT,
            hole_type TEXT,
            hole_class TEXT,
            drilled_by TEXT,
            company TEXT,
            assay_lab TEXT,
            project_id TEXT,
            claim TEXT,
            prospect TEXT,
            purpose TEXT,
            project TEXT,
            survey_type TEXT,
            year TEXT,
            zone TEXT,
            rig_type TEXT,
            rig_no INT,
            core_size TEXT,
            core_storage TEXT,
            remarks TEXT,
            status TEXT,
            dup_id TEXT
        )
    ''')

    # Create tables
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS detailedlog_composite (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hole_id TEXT NOT NULL,
            from_l DECIMAL(7,2),
            to_l DECIMAL(7,2),
            run_l DECIMAL(7,2),
            litho_1 TEXT,
            litho_2 TEXT,
            struc_1 TEXT,
            struc_2 TEXT,
            alt_1 TEXT,
            alt_2 TEXT,
            description TEXT,
            date_relogged TEXT,
            logger TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS lithology_ref (
            litho_1 TEXT,
            litho_2 TEXT,
            remarks TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS alteration_ref (
            alt_1 TEXT,
            alt_2 TEXT,
            remarks TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS structure_ref (
            structure_1 TEXT,
            structure_2 TEXT,
            remarks TEXT
        )
    ''')

    cursor.executemany('''
        INSERT INTO lithology_ref (litho_1, litho_2, remarks) VALUES (?, ?, ?)
    ''', [
        ('Overburden', 'OVER', 'Overburden; transported soil or sidecasted waste'),
        ('Tuff', 'PYCL', 'Volcanic Rock - Tuff, crystal tuff, lithic tuff, lapili tuff, tuff breccia, volcanic breccia, volcaniclastics; andesitic tuff, dacitic tuff embedded in a tuffaceous matrix (fine-grained to crystal-phyric generally with sharp contact with composition from dacite to andesite to basalt).'),
        ('Dacite', 'IDAC', 'Shallow intrusive complex; Dacite; comprises with moderately (40%) abundant phenocrysts of plagioclase and hornblende.'),
        ('Andesite', 'ANFL', 'Volcanic Rock - Andesite; extrusive andesite flow, trachytic andesite; fine-grained to feldspar-phyric porphyritic texture; dense sericite alteration replacement occurs along plagioclase phenocrysts.'),
        ('Diatreme Breccia', 'DBRX', 'Diatreme Breccia - poorly sorted, rounded to subrounded heterolithic clasts of older rock units with varying sizes (millimeter to meters sizes) consisting of stratified but disorientated tuffaceous sedimentary rocks, pyroclastic fallout and breccias, diorite and limestone within a fine ash-rich, tuffaceous sedimentary matrix. '),
        ('Diorite', 'DIO', 'Intrusive Complex; Diorite, Quartz Diorite, and Microdiorite; medium to coarse grained; plagioclase-phyric; pervasively altered to sercite-illite with fine-grained tourmaline. '),
        ('Intrusive Porphyritic Andesite', 'IPYA', 'Intrusive complex (?); massive to moderately abundant phenocrysts in fine-grained groundmass inferred to have formed as hypabyssal dikes.'),
        ('Limestone', 'LMS', 'Sedimentary rock - Limestone described as sequence of bedded calcareous mudstone-siltstone-sandstone-conglomerate intercalated with limestone.'),
        ('BX1', 'BX1', 'Early hydrothermal brecciation event from a deeper diorite intrusion? Composed of silicified and argillized volcanic rocks with subhedral relict phenocrysts cemented by grey quartz, crisscrossed by minimal hairline quartz stockworks.'),
        ('BX2', 'BX2', 'Related to high level dacite intrusions? Composed of the following classes (VNQ, VNB, VNX)'),
        ('BX3', 'BX3', 'Late hydrothermal explosion events after BX2? Multistage brecciation composed of BX1 and BX2 with altered breccia matrix and silica cement.'),
        ('QSX1', 'QSX1', 'BX1 overprinted with mm-cm size veinlets or quartz stockworks.'),
        ('QSW', 'QSW', 'Quartz stockworks mm-cm size in volcanic rocks.'),
        ('VNQ', 'VNQ', 'Massive Quartz Vein - open space infills of milky quartz'),
        ('VNB', 'VNB', 'Banded Quartz Vein - colloform to crustiform banding with notable ginguro bands'),
        ('VNX', 'VNX', 'Quartz Vein Breccia - clasts composed of BX1 and volcanic rocks.'),
        ('CSW', 'CSW', 'Calcite Stockworks'),
        ('CVN', 'CVN', 'Calcite Banded to Massive Veins'),
    ])

    cursor.executemany('''
            INSERT INTO structure_ref (structure_1, structure_2, remarks) VALUES (?, ?, ?)
        ''', [
            ('Fault','FLT','Fault (Normal, Reverse, S-S)'),
            ('Bedding','BED','Bedding'),
            ('Joint', 'JNT', 'Joint'),
            ('Vein', 'VEN', 'Vein'),
            ('Shear', 'SHR', 'Shear Zone'),
            ('Fault Breccia', 'FZB', 'Fault Zone Breccia'),
            ('Lineations', 'LIN', ''),
            ('Crenulation','CRN',''),
            ('Banding', 'BND', 'Banding/Platy Alignment'),
            ('Foliation','FOL',''),
            ('Cavity','CAV','Cavities (previous tunnelway or karst)')
    ])

    cursor.executemany('''
            INSERT INTO alteration_ref (alt_1, alt_2, remarks) VALUES (?, ?, ?)
        ''', [
            ('Advance Argillic','AA','Advance Argillic'),
            ('Argillic', 'AR', 'Argillic'),
            ('Chloritic', 'CH', 'Chloritized'),
            ('Intermiddiate Argillic', 'IA', 'Intermediate Argillic'),
            ('Sericitic', 'SR', 'Sericitic'),
            ('Propylitic', 'PR', 'Propylitic'),
            ('Potassic', 'PT', 'Potassic'),
            ('Silicic', 'SI', 'Silicic'),
            ('Unaltered', 'UA', 'Unaltered'),
            ('Oxidized', 'OX', 'Oxidized Zone'),
            ('Hematitic', 'HM', 'Hematite'),
    ])

    db_connection.commit()

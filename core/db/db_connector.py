import sqlalchemy as sa
from sqlalchemy.orm import sessionmaker, scoped_session
from sqlalchemy import func, inspect, text, and_
from .models import *
from os import getenv
from dotenv import load_dotenv


class Database:
    def __init__(self):
        load_dotenv()
        
        self.db_credentials = "mysql+pymysql://{user}:{password}@{host}:{port}/{db_name}".format(
            user=getenv("DB_USERNAME"),
            password=getenv("DB_PASSWORD"),
            host=getenv("DB_HOSTNAME"),
            port=getenv("DB_PORT"),
            db_name=getenv("DB_NAME")
        )

        self.engine = sa.create_engine(self.db_credentials)
        self.session = scoped_session(sessionmaker(bind=self.engine))

    def get_initial_data_id(
            self,
            project_code,
            pole_code
    ):
        with self.session() as session:
            with session.begin():
                result = session.query(InitialDatas.id)\
                            .filter(and_(
                                InitialDatas.project_code == project_code,
                                InitialDatas.pole_code == pole_code
                            )).scalar()
                print(result)
                return result
    
    def add_initial_data(
            self,
            project_name,
            project_code,
            pole_code,
            pole_type,
            developer,
            voltage,
            area,
            branches,
            wind_region,
            wind_pressure,
            ice_region,
            ice_thickness,
            ice_wind_pressure,
            year_average_temp,
            min_temp,
            max_temp,
            ice_temp,
            wind_temp,
            wind_reg_coef,
            ice_reg_coef,
            wire_hesitation,
            wire,
            wire_tencion,
            ground_wire,
            oksn,
            wind_span,
            weight_span,
            is_stand,
            is_plate,
            txt_1,
            txt_2
        ):
        with self.session() as session:
            with session.begin():
                initial_data_id = session.query(InitialDatas.id)\
                    .filter(and_(
                        InitialDatas.project_code == project_code,
                        InitialDatas.project_name == pole_code
                    )).scalar()
                
                if initial_data_id is None:
                    query = InitialDatas(
                        project_name=project_name,
                        project_code=project_code,
                        pole_code=pole_code,
                        pole_type=pole_type,
                        developer=developer,
                        voltage=voltage,
                        area_type=area,
                        branch=branches,
                        wind_area=wind_region,
                        wind_pressure=wind_pressure,
                        ice_area=ice_region,
                        ice_thickness=ice_thickness,
                        ice_wind_pressure=ice_wind_pressure,
                        avg_temp=year_average_temp,
                        min_temp=min_temp,
                        max_temp=max_temp,
                        ice_temp=ice_temp,
                        wind_temp=wind_temp,
                        wind_reg_coef=wind_reg_coef,
                        ice_reg_coef=ice_reg_coef,
                        wire_hesitation=wire_hesitation,
                        wire=wire,
                        wire_tension=wire_tencion,
                        ground_wire=ground_wire,
                        oksn=oksn,
                        wind_span=wind_span,
                        weight_span=weight_span,
                        is_stand=is_stand,
                        is_plate=is_plate,
                        txt_1=txt_1,
                        txt_2=txt_2
                    )
                    session.add(query)
                else:
                    session.query(InitialDatas)\
                        .where(InitialDatas.id == initial_data_id)\
                            .update(
                                {InitialDatas.project_name: project_name,
                                InitialDatas.project_code: project_code,
                                InitialDatas.pole_code: pole_code,
                                InitialDatas.pole_type: pole_type,
                                InitialDatas.developer: developer,
                                InitialDatas.voltage: voltage,
                                InitialDatas.area_type: area,
                                InitialDatas.branch: branches,
                                InitialDatas.wind_area: wind_region,
                                InitialDatas.wind_pressure: wind_pressure,
                                InitialDatas.ice_area: ice_region,
                                InitialDatas.ice_thickness: ice_thickness,
                                InitialDatas.ice_wind_pressure: ice_wind_pressure,
                                InitialDatas.avg_temp: year_average_temp,
                                InitialDatas.min_temp: min_temp,
                                InitialDatas.max_temp: max_temp,
                                InitialDatas.ice_temp: ice_temp,
                                InitialDatas.wind_temp: wind_temp,
                                InitialDatas.wind_reg_coef: wind_reg_coef,
                                InitialDatas.ice_reg_coef: ice_reg_coef,
                                InitialDatas.wire_hesitation: wire_hesitation,
                                InitialDatas.wire: wire,
                                InitialDatas.wire_tension: wire_tencion,
                                InitialDatas.ground_wire: ground_wire,
                                InitialDatas.oksn: oksn,
                                InitialDatas.wind_span: wind_span,
                                InitialDatas.weight_span: weight_span,
                                InitialDatas.is_stand: is_stand,
                                InitialDatas.is_plate: is_plate,
                                InitialDatas.txt_1: txt_1,
                                InitialDatas.txt_2: txt_2}
                            )

            session.commit()

    def get_initial_data(
            self,
            project_code,
            pole_code
        ):
        with self.session() as session:
            with session.begin():
                query = session.query(InitialDatas).filter(
                    and_(
                        InitialDatas.project_code == project_code,
                        InitialDatas.pole_code == pole_code
                    )
                ).scalar()
                if query:
                    result = {
                        "initial_data_id": query.id,
                        "project_name": query.project_name,
                        "project_code": query.project_code,
                        "pole_code": query.pole_code,
                        "pole_type": query.pole_type,
                        "developer": query.developer,
                        "voltage": query.voltage,
                        "area_type": query.area_type,
                        "branch": query.branch,
                        'wind_area': query.wind_area,
                        "wind_pressure": query.wind_pressure,
                        "ice_area": query.ice_area,
                        "ice_thickness": query.ice_thickness,
                        "ice_wind_pressure": query.ice_wind_pressure,
                        "avg_temp": query.avg_temp,
                        "min_temp": query.min_temp,
                        "max_temp": query.max_temp,
                        "ice_temp": query.ice_temp,
                        "wind_temp": query.wind_temp,
                        "wind_reg_coef": query.wind_reg_coef,
                        "ice_reg_coef": query.ice_reg_coef,
                        "wire_hesitation": query.wire_hesitation,
                        "wire": query.wire,
                        "wire_tension": query.wire_tension,
                        "ground_wire": query.ground_wire,
                        "oksn": query.oksn,
                        "wind_span": query.wind_span,
                        "weight_span": query.weight_span,
                        "is_stand": query.is_stand,
                        "is_plate": query.is_plate,
                        "txt_1": query.txt_1,
                        "txt_2": query.txt_2
                    }
                    return result
                else:
                    return None
    
    def add_rpzo_data(
        self,
        wire_pos,
        ground_wire_attachment,
        quantity_of_ground_wire,
        is_ground_wire_davit,
        deflection,
        pls_pole_pic1,
        pls_pole_pic2,
        loads_1,
        loads_2,
        loads_3,
        loads_4,
        loads_5,
        loads_6,
        is_mont_schema,
        initial_data_id
    ):
        with self.session() as session:
            with session.begin():
                rpzo_data = session.query(RpzoDatas)\
                    .filter(RpzoDatas.initial_data_id == initial_data_id)\
                        .scalar()
                
                if rpzo_data is None:
                    query = RpzoDatas(
                        wire_pos=wire_pos,
                        ground_wire_attachment=ground_wire_attachment,
                        quantity_of_ground_wire=quantity_of_ground_wire,
                        is_ground_wire_davit=is_ground_wire_davit,
                        deflection=deflection,
                        pls_pole_pic1=pls_pole_pic1,
                        pls_pole_pic2=pls_pole_pic2,
                        loads_1=loads_1,
                        loads_2=loads_2,
                        loads_3=loads_3,
                        loads_4=loads_4,
                        loads_5=loads_5,
                        loads_6=loads_6,
                        is_mont_schema=is_mont_schema,
                        initial_data_id=initial_data_id
                    )
                    session.add(query)
                else:
                    session.query(RpzoDatas)\
                        .where(RpzoDatas.initial_data_id == initial_data_id)\
                            .update(
                                {RpzoDatas.wire_pos: wire_pos,
                                RpzoDatas.ground_wire_attachment: ground_wire_attachment,
                                RpzoDatas.quantity_of_ground_wire: quantity_of_ground_wire,
                                RpzoDatas.is_ground_wire_davit: is_ground_wire_davit,
                                RpzoDatas.deflection: deflection,
                                RpzoDatas.pls_pole_pic1: pls_pole_pic1,
                                RpzoDatas.pls_pole_pic2: pls_pole_pic2,
                                RpzoDatas.loads_1: loads_1,
                                RpzoDatas.loads_2: loads_2,
                                RpzoDatas.loads_3: loads_3,
                                RpzoDatas.loads_4: loads_4,
                                RpzoDatas.loads_5: loads_5,
                                RpzoDatas.loads_6: loads_6,
                                RpzoDatas.is_mont_schema: is_mont_schema,
                                RpzoDatas.initial_data_id: initial_data_id}
                            )
                session.commit()

    def get_rpzo_data(self, initial_data_id):
        with self.session() as session:
            with session.begin():
                query = session.query(RpzoDatas)\
                    .filter(RpzoDatas.initial_data_id == initial_data_id)\
                        .scalar()
                if query:
                    result = {
                        "wire_pos": query.wire_pos,
                        "ground_wire_attachment": query.ground_wire_attachment,
                        "quantity_of_ground_wire": query.quantity_of_ground_wire,
                        "is_ground_wire_davit": query.is_ground_wire_davit,
                        "deflection": query.deflection,
                        "pls_pole_pic1": query.pls_pole_pic1,
                        "pls_pole_pic2": query.pls_pole_pic2,
                        "loads_1": query.loads_1,
                        "loads_2": query.loads_2,
                        "loads_3": query.loads_3,
                        "loads_4": query.loads_4,
                        "loads_5": query.loads_5,
                        "loads_6": query.loads_6,
                        "is_mont_schema": query.is_mont_schema
                    }
                    return result
                else:
                    return None
                
    def add_foundation_data(
            self,
            initial_data_id,
            thickness_svai,
            deepness_svai,
            height_svai,
            is_initial_data,
            typical_ground,
            udel_sceplenie,
            ugol_vntr_trenia,
            ves_grunta,
            deform_module,
            ground_water_lvl,
            quantity_of_sloy,
            nomer_ige1,
            nomer_ige2,
            nomer_ige3,
            nomer_ige4,
            nomer_ige5,
            ground_type1,
            ground_type2,
            ground_type3,
            ground_type4,
            ground_type5,
            ground_name1,
            ground_name2,
            ground_name3,
            ground_name4,
            ground_name5,
            verh_sloy1,
            verh_sloy2,
            verh_sloy3,
            verh_sloy4,
            verh_sloy5,
            nijn_sloy1,
            nijn_sloy2,
            nijn_sloy3,
            nijn_sloy4,
            nijn_sloy5,
            coef_poristosti1,
            coef_poristosti2,
            coef_poristosti3,
            coef_poristosti4,
            coef_poristosti5,
            udel_scep1,
            udel_scep2,
            udel_scep3,
            udel_scep4,
            udel_scep5,
            ugol_vn_tr1,
            ugol_vn_tr2,
            ugol_vn_tr3,
            ugol_vn_tr4,
            ugol_vn_tr5,
            ves_gr_prir1,
            ves_gr_prir2,
            ves_gr_prir3,
            ves_gr_prir4,
            ves_gr_prir5,
            def_mod1,
            def_mod2,
            def_mod3,
            def_mod4,
            def_mod5,
            pole_type,
            coef_nadej
    ):
        with self.session() as session:
            with session.begin():
                foundation_data = session.query(FoundationDatas)\
                    .filter(FoundationDatas.initial_data_id == initial_data_id)\
                        .scalar()
                
                if foundation_data is None:
                    query = FoundationDatas(
                        initial_data_id=initial_data_id,
                        thickness_svai=thickness_svai,
                        deepness_svai=deepness_svai,
                        height_svai=height_svai,
                        is_initial_data=is_initial_data,
                        typical_ground=typical_ground,
                        udel_sceplenie=udel_sceplenie,
                        ugol_vntr_trenia=ugol_vntr_trenia,
                        ves_grunta=ves_grunta,
                        deform_module=deform_module,
                        ground_water_lvl=ground_water_lvl,
                        quantity_of_sloy=quantity_of_sloy,
                        nomer_ige1=nomer_ige1,
                        nomer_ige2=nomer_ige2,
                        nomer_ige3=nomer_ige3,
                        nomer_ige4=nomer_ige4,
                        nomer_ige5=nomer_ige5,
                        ground_type1=ground_type1,
                        ground_type2=ground_type2,
                        ground_type3=ground_type3,
                        ground_type4=ground_type4,
                        ground_type5=ground_type5,
                        ground_name1=ground_name1,
                        ground_name2=ground_name2,
                        ground_name3=ground_name3,
                        ground_name4=ground_name4,
                        ground_name5=ground_name5,
                        verh_sloy1=verh_sloy1,
                        verh_sloy2=verh_sloy2,
                        verh_sloy3=verh_sloy3,
                        verh_sloy4=verh_sloy4,
                        verh_sloy5=verh_sloy5,
                        nijn_sloy1=nijn_sloy1,
                        nijn_sloy2=nijn_sloy2,
                        nijn_sloy3=nijn_sloy3,
                        nijn_sloy4=nijn_sloy4,
                        nijn_sloy5=nijn_sloy5,
                        coef_poristosti1=coef_poristosti1,
                        coef_poristosti2=coef_poristosti2,
                        coef_poristosti3=coef_poristosti3,
                        coef_poristosti4=coef_poristosti4,
                        coef_poristosti5=coef_poristosti5,
                        udel_scep1=udel_scep1,
                        udel_scep2=udel_scep2,
                        udel_scep3=udel_scep3,
                        udel_scep4=udel_scep4,
                        udel_scep5=udel_scep5,
                        ugol_vn_tr1=ugol_vn_tr1,
                        ugol_vn_tr2=ugol_vn_tr2,
                        ugol_vn_tr3=ugol_vn_tr3,
                        ugol_vn_tr4=ugol_vn_tr4,
                        ugol_vn_tr5=ugol_vn_tr5,
                        ves_gr_prir1=ves_gr_prir1,
                        ves_gr_prir2=ves_gr_prir2,
                        ves_gr_prir3=ves_gr_prir3,
                        ves_gr_prir4=ves_gr_prir4,
                        ves_gr_prir5=ves_gr_prir5,
                        def_mod1=def_mod1,
                        def_mod2=def_mod2,
                        def_mod3=def_mod3,
                        def_mod4=def_mod4,
                        def_mod5=def_mod5,
                        pole_type=pole_type,
                        coef_nadej=coef_nadej
                    )
                    session.add(query)
                else:
                    session.query(FoundationDatas)\
                        .filter(FoundationDatas.initial_data_id == initial_data_id)\
                            .update(
                                {FoundationDatas.initial_data_id: initial_data_id,
                                FoundationDatas.thickness_svai: thickness_svai,
                                FoundationDatas.deepness_svai: deepness_svai,
                                FoundationDatas.height_svai: height_svai,
                                FoundationDatas.is_initial_data: is_initial_data,
                                FoundationDatas.typical_ground: typical_ground,
                                FoundationDatas.udel_sceplenie: udel_sceplenie,
                                FoundationDatas.ugol_vntr_trenia: ugol_vntr_trenia,
                                FoundationDatas.ves_grunta: ves_grunta,
                                FoundationDatas.deform_module: deform_module,
                                FoundationDatas.ground_water_lvl: ground_water_lvl,
                                FoundationDatas.quantity_of_sloy: quantity_of_sloy,
                                FoundationDatas.nomer_ige1: nomer_ige1,
                                FoundationDatas.nomer_ige2: nomer_ige2,
                                FoundationDatas.nomer_ige3: nomer_ige3,
                                FoundationDatas.nomer_ige4: nomer_ige4,
                                FoundationDatas.nomer_ige5: nomer_ige5,
                                FoundationDatas.ground_type1: ground_type1,
                                FoundationDatas.ground_type2: ground_type2,
                                FoundationDatas.ground_type3: ground_type3,
                                FoundationDatas.ground_type4: ground_type4,
                                FoundationDatas.ground_type5: ground_type5,
                                FoundationDatas.ground_name1: ground_name1,
                                FoundationDatas.ground_name2: ground_name2,
                                FoundationDatas.ground_name3: ground_name3,
                                FoundationDatas.ground_name4: ground_name4,
                                FoundationDatas.ground_name5: ground_name5,
                                FoundationDatas.verh_sloy1: verh_sloy1,
                                FoundationDatas.verh_sloy2: verh_sloy2,
                                FoundationDatas.verh_sloy3: verh_sloy3,
                                FoundationDatas.verh_sloy4: verh_sloy4,
                                FoundationDatas.verh_sloy5: verh_sloy5,
                                FoundationDatas.nijn_sloy1: nijn_sloy1,
                                FoundationDatas.nijn_sloy2: nijn_sloy2,
                                FoundationDatas.nijn_sloy3: nijn_sloy3,
                                FoundationDatas.nijn_sloy4: nijn_sloy4,
                                FoundationDatas.nijn_sloy5: nijn_sloy5,
                                FoundationDatas.coef_poristosti1: coef_poristosti1,
                                FoundationDatas.coef_poristosti2: coef_poristosti2,
                                FoundationDatas.coef_poristosti3: coef_poristosti3,
                                FoundationDatas.coef_poristosti4: coef_poristosti4,
                                FoundationDatas.coef_poristosti5: coef_poristosti5,
                                FoundationDatas.udel_scep1: udel_scep1,
                                FoundationDatas.udel_scep2: udel_scep2,
                                FoundationDatas.udel_scep3: udel_scep3,
                                FoundationDatas.udel_scep4: udel_scep4,
                                FoundationDatas.udel_scep5: udel_scep5,
                                FoundationDatas.ugol_vn_tr1: ugol_vn_tr1,
                                FoundationDatas.ugol_vn_tr2: ugol_vn_tr2,
                                FoundationDatas.ugol_vn_tr3: ugol_vn_tr3,
                                FoundationDatas.ugol_vn_tr4: ugol_vn_tr4,
                                FoundationDatas.ugol_vn_tr5: ugol_vn_tr5,
                                FoundationDatas.ves_gr_prir1: ves_gr_prir1,
                                FoundationDatas.ves_gr_prir2: ves_gr_prir2,
                                FoundationDatas.ves_gr_prir3: ves_gr_prir3,
                                FoundationDatas.ves_gr_prir4: ves_gr_prir4,
                                FoundationDatas.ves_gr_prir5: ves_gr_prir5,
                                FoundationDatas.def_mod1: def_mod1,
                                FoundationDatas.def_mod2: def_mod2,
                                FoundationDatas.def_mod3: def_mod3,
                                FoundationDatas.def_mod4: def_mod4,
                                FoundationDatas.def_mod5: def_mod5,
                                FoundationDatas.pole_type: pole_type,
                                FoundationDatas.coef_nadej: coef_nadej}
                            )
                session.commit()
    
    def get_foundation_data(self, initial_data_id):
        with self.session() as session:
            with session.begin():
                query = session.query(FoundationDatas)\
                    .filter(FoundationDatas.initial_data_id == initial_data_id)\
                        .scalar()
                if query:
                    result = {
                        'initial_data_id': query.initial_data_id,
                        'thickness_svai': query.thickness_svai,
                        'deepness_svai': query.deepness_svai,
                        'height_svai': query.height_svai,
                        'is_initial_data': query.is_initial_data,
                        'typical_ground': query.typical_ground,
                        'udel_sceplenie': query.udel_sceplenie,
                        'ugol_vntr_trenia': query.ugol_vntr_trenia,
                        'ves_grunta': query.ves_grunta,
                        'deform_module': query.deform_module,
                        'ground_water_lvl': query.ground_water_lvl,
                        'quantity_of_sloy': query.quantity_of_sloy,
                        'nomer_ige1': query.nomer_ige1,
                        'nomer_ige2': query.nomer_ige2,
                        'nomer_ige3': query.nomer_ige3,
                        'nomer_ige4': query.nomer_ige4,
                        'nomer_ige5': query.nomer_ige5,
                        'ground_type1': query.ground_type1,
                        'ground_type2': query.ground_type2,
                        'ground_type3': query.ground_type3,
                        'ground_type4': query.ground_type4,
                        'ground_type5': query.ground_type5,
                        'ground_name1': query.ground_name1,
                        'ground_name2': query.ground_name2,
                        'ground_name3': query.ground_name3,
                        'ground_name4': query.ground_name4,
                        'ground_name5': query.ground_name5,
                        'verh_sloy1': query.verh_sloy1,
                        'verh_sloy2': query.verh_sloy2,
                        'verh_sloy3': query.verh_sloy3,
                        'verh_sloy4': query.verh_sloy4,
                        'verh_sloy5': query.verh_sloy5,
                        'nijn_sloy1': query.nijn_sloy1,
                        'nijn_sloy2': query.nijn_sloy2,
                        'nijn_sloy3': query.nijn_sloy3,
                        'nijn_sloy4': query.nijn_sloy4,
                        'nijn_sloy5': query.nijn_sloy5,
                        'coef_poristosti1': query.coef_poristosti1,
                        'coef_poristosti2': query.coef_poristosti2,
                        'coef_poristosti3': query.coef_poristosti3,
                        'coef_poristosti4': query.coef_poristosti4,
                        'coef_poristosti5': query.coef_poristosti5,
                        'udel_scep1': query.udel_scep1,
                        'udel_scep2': query.udel_scep2,
                        'udel_scep3': query.udel_scep3,
                        'udel_scep4': query.udel_scep4,
                        'udel_scep5': query.udel_scep5,
                        'ugol_vn_tr1': query.ugol_vn_tr1,
                        'ugol_vn_tr2': query.ugol_vn_tr2,
                        'ugol_vn_tr3': query.ugol_vn_tr3,
                        'ugol_vn_tr4': query.ugol_vn_tr4,
                        'ugol_vn_tr5': query.ugol_vn_tr5,
                        'ves_gr_prir1': query.ves_gr_prir1,
                        'ves_gr_prir2': query.ves_gr_prir2,
                        'ves_gr_prir3': query.ves_gr_prir3,
                        'ves_gr_prir4': query.ves_gr_prir4,
                        'ves_gr_prir5': query.ves_gr_prir5,
                        'def_mod1': query.def_mod1,
                        'def_mod2': query.def_mod2,
                        'def_mod3': query.def_mod3,
                        'def_mod4': query.def_mod4,
                        'def_mod5': query.def_mod5,
                        'pole_type': query.pole_type,
                        'coef_nadej': query.coef_nadej
                    }
                    return result
                else:
                    return None
                
    def add_rpzf_data(
            self,
            initial_data_id,
            ige_name,
            building_adress,
            razrez_skvajin,
            picture1,
            picture2,
            xlsx_svai
    ):
        with self.session() as session:
            with session.begin():
                rpzf_data = session.query(RpzfDatas)\
                    .filter(RpzfDatas.initial_data_id == initial_data_id)\
                        .scalar()
                if rpzf_data is None:
                    query = RpzfDatas(
                        initial_data_id=initial_data_id,
                        ige_name=ige_name,
                        building_adress=building_adress,
                        razrez_skvajin=razrez_skvajin,
                        picture1=picture1,
                        picture2=picture2,
                        xlsx_svai=xlsx_svai
                    )
                    session.add(query)
                else:
                    session.query(RpzfDatas)\
                        .filter(RpzfDatas.initial_data_id == initial_data_id)\
                            .update(
                                {
                                    RpzfDatas.ige_name: ige_name,
                                    RpzfDatas.building_adress: building_adress,
                                    RpzfDatas.razrez_skvajin: razrez_skvajin,
                                    RpzfDatas.picture1: picture1,
                                    RpzfDatas.picture2: picture2,
                                    RpzfDatas.xlsx_svai: xlsx_svai
                                }
                            )
                session.commit()

    def get_rpzf_data(self, initial_data_id):
        with self.session() as session:
            with session.begin():
                query = session.query(RpzfDatas)\
                    .filter(RpzfDatas.initial_data_id == initial_data_id)\
                        .scalar()
                if query:
                    result = {
                        'initial_data_id': query.initial_data_id,
                        'ige_name': query.ige_name,
                        'building_adress': query.building_adress,
                        'razrez_skvajin': query.razrez_skvajin,
                        'picture1': query.picture1,
                        'picture2': query.picture2,
                        'xlsx_svai': query.xlsx_svai
                    }
                    return result
                else:
                    return None
                
    def get_svai_data(self, initial_data_id):
        with self.session() as session:
            with session.begin():
                query = session.query(FoundationDatas)\
                    .filter(FoundationDatas.initial_data_id == initial_data_id)\
                        .scalar()
                result = {
                    "deepness_svai": query.deepness_svai,
                    "height_svai": query.height_svai
                }
                return result
            
    def add_anker_data(
            self,
            initial_data_id,
            bolt,
            kol_boltov,
            bolt_class,
            hole_diam,
            rast_m,
            xlsx_bolt,
            bolt_schema
    ):
        with self.session() as session:
            with session.begin():
                anker_data = session.query(AnkerDatas)\
                    .filter(AnkerDatas.initial_data_id == initial_data_id)\
                        .scalar()
                if anker_data is None:
                    query = AnkerDatas(
                        initial_data_id=initial_data_id,
                        bolt=bolt,
                        kol_boltov=kol_boltov,
                        bolt_class=bolt_class,
                        hole_diam=hole_diam,
                        rast_m=rast_m,
                        xlsx_bolt=xlsx_bolt,
                        bolt_schema=bolt_schema
                    )
                    session.add(query)
                else:
                    session.query(AnkerDatas)\
                        .filter(AnkerDatas.initial_data_id == initial_data_id)\
                            .update(
                                {
                                    AnkerDatas.bolt: bolt,
                                    AnkerDatas.kol_boltov: kol_boltov,
                                    AnkerDatas.bolt_class: bolt_class,
                                    AnkerDatas.hole_diam: hole_diam,
                                    AnkerDatas.rast_m: rast_m,
                                    AnkerDatas.xlsx_bolt: xlsx_bolt,
                                    AnkerDatas.bolt_schema: bolt_schema
                                }
                            )
                session.commit()

    def get_anker_data(self, initial_data_id):
        with self.session() as session:
            with session.begin():
                query = session.query(AnkerDatas)\
                    .filter(AnkerDatas.initial_data_id == initial_data_id)\
                        .scalar()
                if query:
                    result = {
                        "bolt": query.bolt,
                        "kol_boltov": query.kol_boltov,
                        "bolt_class": query.bolt_class,
                        "hole_diam": query.hole_diam,
                        "rast_m": query.rast_m,
                        "xlsx_bolt": query.xlsx_bolt,
                        "bolt_schema": query.bolt_schema
                    }
                    return result
                else:
                    return None
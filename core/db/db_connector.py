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
                return session.query(InitialDatas.id)\
                            .filter(and_(
                                InitialDatas.project_code == project_code,
                                InitialDatas.project_name == pole_code
                            )).scalar()
    
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
                initial_data_id = self.get_initial_data(
                    project_code=project_code,
                    pole_code=pole_code
                )
                
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

    ):
        pass
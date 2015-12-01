from sqlalchemy import Column, Integer, String, Unicode, Boolean, create_engine, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

__all__ = ['Hesaplar', 'Lead', 'session', 'tanimlar']

Base = declarative_base()

db_name = ':memory:'
tanimlar = {'company': 'Fatih Ka.'}
periods = dict(cy='31.12.2015', py1='31.12.2013', py2='31.12.2014')


class Hesaplar(Base):
    __tablename__ = 'hesaplar'

    id = Column(Integer, primary_key=True)
    number = Column(String, nullable=True)
    ana_hesap = Column(String, nullable=True)
    name = Column(Unicode, nullable=True)
    lead_code = Column(String, default="Unmapped", nullable=True)
    cy = Column(Float, nullable=True, default=0)
    py1 = Column(Float, nullable=True, default=0)
    py2 = Column(Float, nullable=True, default=0)
    len = Column(Integer, nullable=True)
    bd = Column(Boolean, nullable=False, default=False)


class Lead(Base):
    __tablename__ = 'ana_hesaplar'

    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=True)
    lead_code = Column(String, nullable=True)
    account = Column(String, nullable=True)


engine = create_engine("sqlite:///%s" % db_name, echo=False)

Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session = Session()

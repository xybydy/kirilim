__author__ = 'fatihka'

from sqlalchemy import Column, Integer, String, Unicode, Float, Boolean, create_engine, Table
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

__all__ = ['Hesaplar', 'Lead', 'session', 'tanimlar']

Base = declarative_base()

db_name = 'qq.db'
tanimlar = {'company': 'Fatih Ka.'}
periodss = list()

# class Hesaplar(Base):
#     __tablename__ = 'hesaplar'
#
#     id = Column(Integer, primary_key=True)
#     number = Column(String, nullable=True)
#     ana_hesap = Column(String, nullable=True)
#     name = Column(Unicode, nullable=True)
#     lead_code = Column(String, default="Unmapped", nullable=True)
#     cy = Column(Float, nullable=True, default=0)
#     py1 = Column(Float, nullable=True, default=0)
#     py2 = Column(Float, nullable=True, default=0)
#     len = Column(Integer, nullable=True)
#     bd = Column(Boolean, nullable=False, default=False)

Hesaplar = None
session = None


class Lead(Base):
    __tablename__ = 'ana_hesaplar'
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=True)
    lead_code = Column(String, nullable=True)
    account = Column(String, nullable=True)
    account_name = Column(String, nullable=True)


def make_hesaplar():
    class Hesaplar(Base):
        __table__ = Table('hesaplar', Base.metadata,
                          Column('id', Integer, primary_key=True),
                          Column('number', String, nullable=True),
                          Column('ana_hesap', String, nullable=True),
                          Column('name', Unicode, nullable=True),
                          Column('lead_code', String, default='Unmapped', nullable=True),
                          Column('len', Integer, nullable=True),
                          Column('bd', Boolean, nullable=True, default=False),
                          *[Column('%s'%i, Float, nullable=True, default=0) for i in periodss]
                          )

    return Hesaplar


def create_db():
    global session
    engine = create_engine("sqlite:///%s" % db_name, echo=False)  # engine = create_engine("sqlite://", echo=False)
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)
    session = Session()




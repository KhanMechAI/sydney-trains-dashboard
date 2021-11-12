import os
import sys
from sqlalchemy import Column, ForeignKey, Integer, String, DateTime
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from sqlalchemy import create_engine

Base = declarative_base()

DB_ENGINE_NAME = 'sqlite:///key_client_database.db'

class ClientOrganisation(Base):
    __tablename__ = "client_organisation"
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    project_manager = relationship("ClientProjectManager", back_populates="client_organisation")

class ClientProjectManager(Base):
    __tablename__ = "client_project_managers"
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    company_id = Column(Integer, ForeignKey("client_organisation.id"))
    company = relationship("ClientOrganisation", uselist=False, back_populates="client_project_managers")

class GHDProjectManager(Base):
    __tablename__ = "ghd_project_managers"
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    projects = relationship("Projects")


class Projects(Base):
    __tablename__ = "projects"
    id = Column(Integer, primary_key=True)
    project_number = Column(Integer, nullable=False)
    client_po = Column(Integer)
    client_project_number = Column(String)
    project_title = Column(String(250), nullable=False)
    project_manager = Column(Integer, ForeignKey("ghd_project_managers.id"))
    project_manager = relationship("GHDProjectManager", back_populates="ghd_project_managers")
    client_project_manager = Column(Integer, ForeignKey("client_project_managers.id"))
    client_project_manager = relationship("ClientProjectManager", back_populates="client_project_managers")
    phase = Column(String(100))
    schedule = Column(String(25))
    contractual_completion_date = Column(String(25))
    current_status = Column(String(250))
    next_actions = Column(String(250))
    action_by = Column(String(50))


engine = create_engine(DB_ENGINE_NAME)

Base.metadata.create_all(engine)
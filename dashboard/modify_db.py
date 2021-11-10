from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from db import Projects, GHDProjectManager, ClientProjectManager, ClientOrganisation
from db import DB_ENGINE_NAME

from dashboard import Dashboard, BSTLoader

engine = create_engine(DB_ENGINE_NAME)

DBSession = sessionmaker(bind=engine)

session = DBSession()





dash = Dashboard()
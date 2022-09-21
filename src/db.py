from sqlalchemy import create_engine
from functools import lru_cache
from contextlib import contextmanager


@lru_cache(maxsize=1)
def init_db_connection(config):
    url = f"{config['protocol']}://{config['tns']}/?encoding={config['encoding']}"
    engine = create_engine(url)
    return engine


@contextmanager
def get_connection_context(config):
    engine = init_db_connection(config)
    with engine.connect() as conn:
        yield conn

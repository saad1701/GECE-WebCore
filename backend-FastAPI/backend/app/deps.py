
# Shared dependencies
from fastapi import Depends, Header

def get_actor(x_actor: str | None = Header(default='system')):
    return x_actor or 'system'

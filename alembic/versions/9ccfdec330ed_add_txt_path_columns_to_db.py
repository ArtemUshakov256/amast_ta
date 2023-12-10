"""add txt path columns to db

Revision ID: 9ccfdec330ed
Revises: be1b075b6eef
Create Date: 2023-11-16 16:56:37.496942

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = '9ccfdec330ed'
down_revision: Union[str, None] = 'be1b075b6eef'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # ### commands auto generated by Alembic - please adjust! ###
    op.add_column('initial_datas', sa.Column('txt_1', sa.Text(), nullable=False))
    op.add_column('initial_datas', sa.Column('txt_2', sa.Text(), nullable=False))
    # ### end Alembic commands ###


def downgrade() -> None:
    # ### commands auto generated by Alembic - please adjust! ###
    op.drop_column('initial_datas', 'txt_2')
    op.drop_column('initial_datas', 'txt_1')
    # ### end Alembic commands ###
"""remove some fields

Revision ID: 004f1500e67c
Revises: c27f222e9256
Create Date: 2023-11-18 15:27:59.525119

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import mysql

# revision identifiers, used by Alembic.
revision: str = '004f1500e67c'
down_revision: Union[str, None] = 'c27f222e9256'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # ### commands auto generated by Alembic - please adjust! ###
    op.drop_column('foundation_datas', 'vert_force')
    op.drop_column('foundation_datas', 'shear_force')
    op.drop_column('foundation_datas', 'bending_moment')
    op.drop_column('foundation_datas', 'diam_svai')
    # ### end Alembic commands ###


def downgrade() -> None:
    # ### commands auto generated by Alembic - please adjust! ###
    op.add_column('foundation_datas', sa.Column('diam_svai', mysql.VARCHAR(length=5), nullable=False))
    op.add_column('foundation_datas', sa.Column('bending_moment', mysql.VARCHAR(length=10), nullable=False))
    op.add_column('foundation_datas', sa.Column('shear_force', mysql.VARCHAR(length=10), nullable=False))
    op.add_column('foundation_datas', sa.Column('vert_force', mysql.VARCHAR(length=10), nullable=False))
    # ### end Alembic commands ###

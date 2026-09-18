from datetime import date

from sqlalchemy import (
    Boolean,
    Date,
    ForeignKey,
    Integer,
    String,
    Text,
)

from sqlalchemy.orm import Mapped, mapped_column, relationship

from .db import Base


# =========================================================
# CRONOGRAMA
# =========================================================

class Cronograma(Base):
    __tablename__ = "cronogramas"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
        index=True,
    )

    nombre_archivo: Mapped[str] = mapped_column(
        String(255),
        nullable=False,
    )

    recorridos: Mapped[list["Recorrido"]] = relationship(
        "Recorrido",
        back_populates="cronograma",
        cascade="all, delete-orphan",
    )


# =========================================================
# RECORRIDO
# =========================================================

class Recorrido(Base):
    __tablename__ = "recorridos"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
        index=True,
    )

    cronograma_id: Mapped[int] = mapped_column(
        ForeignKey("cronogramas.id"),
        nullable=False,
    )

    funcionario: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    dia: Mapped[str | None] = mapped_column(
        String(50),
        nullable=True,
    )

    fecha: Mapped[date | None] = mapped_column(
        Date,
        nullable=True,
    )

    hora_programada: Mapped[str | None] = mapped_column(
        String(100),
        nullable=True,
    )

    provincia: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    canton: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    parroquia: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    sector: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    calle_inicial: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    calle_final: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    observaciones: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    hay_articulos: Mapped[bool] = mapped_column(
        Boolean,
        default=False,
        nullable=False,
    )

    cronograma: Mapped["Cronograma"] = relationship(
        "Cronograma",
        back_populates="recorridos",
    )

    anexos: Mapped[list["Anexo"]] = relationship(
        "Anexo",
        back_populates="recorrido",
        cascade="all, delete-orphan",
    )


# =========================================================
# ANEXO
# =========================================================

class Anexo(Base):
    __tablename__ = "anexos"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True,
        index=True,
    )

    recorrido_id: Mapped[int] = mapped_column(
        ForeignKey("recorridos.id"),
        nullable=False,
    )

    # -----------------------------------------------------
    # UBICACIÓN
    # -----------------------------------------------------

    circunscripcion: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    ciudad: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    zona: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    direccion: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    referencia: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    # -----------------------------------------------------
    # INFORMACIÓN ELECTORAL / EVIDENCIA
    # -----------------------------------------------------

    organizacion_politica: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    dignidad: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    candidato: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    leyenda: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    articulo: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    medidas: Mapped[str | None] = mapped_column(
        String(100),
        nullable=True,
    )

    placa: Mapped[str | None] = mapped_column(
        String(100),
        nullable=True,
    )

    hora: Mapped[str | None] = mapped_column(
        String(100),
        nullable=True,
    )

    registro: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    cantidad: Mapped[int | None] = mapped_column(
        Integer,
        default=1,
        nullable=True,
    )

    foto_path: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    # -----------------------------------------------------
    # RELACIÓN CON RECORRIDO
    # -----------------------------------------------------

    recorrido: Mapped["Recorrido"] = relationship(
        "Recorrido",
        back_populates="anexos",
    )
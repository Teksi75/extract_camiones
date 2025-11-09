from dataclasses import dataclass
from typing import Optional, List, Dict

@dataclass
class ModeloInstrumento:
    modelo: str
    marca: str
    fabricante: str
    origen: str
    n_aprob: str
    fecha_aprob: str
    tipo_instr: str
    max: str
    min: str
    e: str
    dd_dt: str
    clase: str
    codigo_aprobacion: str

@dataclass
class Instrumento:
    id: str
    domicilio: str
    localidad: str
    provincia: str
    receptor: ModeloInstrumento
    indicador: ModeloInstrumento

@dataclass
class Tramite:
    ot: str
    vpe: str
    empresa: str
    propietario: str
    direccion_fiscal: str
    instrumentos: List[Instrumento]

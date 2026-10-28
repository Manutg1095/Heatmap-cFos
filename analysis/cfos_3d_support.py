"""Código del análisis final de c-Fos; véase analysis/README.md para su alcance."""
import argparse
from pathlib import Path
import sys, numpy as np, pandas as pd
from read_xlsx_xml import XLSX


def run(input_dir, output_dir):
    ROOT=Path(input_dir).expanduser().resolve()
    OUT=Path(output_dir).expanduser().resolve()
    OUT.mkdir(parents=True, exist_ok=True)
    x=XLSX(ROOT/'resultados_cfos_2D_3D.xlsx');q=x.read('QC_y_metricas_stack');cells=x.read('Celulas_3D')
    for d in [q,cells]:d['Animal_ID']=d.Experimento.astype(str)+'::'+d.Animal.astype(str)
    for c in ['Bregma_mm','Espesor_derivado_um']:q[c]=pd.to_numeric(q[c],errors='coerce')
    for c in ['X_ML_norm_0_medial_1_lateral','Y_DV_norm_0_ventral_1_dorsal','AP_mm']:cells[c]=pd.to_numeric(cells[c],errors='coerce')
    XC='X_ML_norm_0_medial_1_lateral';YC='Y_DV_norm_0_ventral_1_dorsal'
    rows=[]
    for f in sorted((ROOT/'matrices_3D').glob('3D_*.npz')):
     a=np.load(f);D=a['densidad_celulas_mm3'];N=a['conteos'];V=a['volumen_um3'];idx=np.unravel_index(np.nanargmax(D),D.shape)
     group=f.stem[len('3D_'):].rsplit('_',1)[0].replace('Veh_culo','Vehículo');reg=f.stem.rsplit('_',1)[1]
     xe,ye,ae=a['x_edges'],a['y_edges'],a['ap_edges_mm'];xl,xh=xe[idx[0]],xe[idx[0]+1];yl,yh=ye[idx[1]],ye[idx[1]+1];al,ah=ae[idx[2]],ae[idx[2]+1]
     cc=cells[(cells.Grupo==group)&(cells.Region==reg)];inside=cc[(cc[XC]>=xl)&(cc[XC]<xh)&(cc[YC]>=yl)&(cc[YC]<yh)&(cc.AP_mm>=al)&(cc.AP_mm<ah)]
     qq=q[(q.Grupo==group)&(q.Region==reg)].copy();qq['low']=qq.Bregma_mm-qq.Espesor_derivado_um/2000;qq['high']=qq.Bregma_mm+qq.Espesor_derivado_um/2000
     va=qq[(qq.high>al)&(qq.low<ah)].Animal_ID.nunique()
     rows.append(dict(Region=reg,Grupo=group,Pico_mm3=float(D[idx]),Celulas_en_voxel=int(N[idx]),Volumen_um3=float(V[idx]),Animales_con_celulas=inside.Animal_ID.nunique(),Animales_con_volumen=int(va),X_centro=(xl+xh)/2,Y_centro=(yl+yh)/2,AP_centro=(al+ah)/2,Animal_objetos=';'.join(sorted(inside.Animal_ID.unique()))))
    out=pd.DataFrame(rows).sort_values(['Region','Grupo']);out.to_csv(OUT/'maximos_3D.csv',index=False);print(out.to_string(index=False))


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True, type=Path, help='Corrida con matrices_3D y libro de resultados')
    parser.add_argument('--output', required=True, type=Path, help='Carpeta de resultados estadísticos')
    args=parser.parse_args()
    run(args.input, args.output)


if __name__ == "__main__":
    main()

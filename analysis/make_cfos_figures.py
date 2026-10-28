"""Código del análisis final de c-Fos; véase analysis/README.md para su alcance."""
import argparse
from pathlib import Path
import sys
import numpy as np,pandas as pd
import matplotlib.pyplot as plt
from scipy.ndimage import gaussian_filter
from read_xlsx_xml import XLSX


def run(input_dir, results_dir, output_dir):
    ROOT=Path(input_dir).expanduser().resolve()
    R=Path(results_dir).expanduser().resolve()
    OUT=Path(output_dir).expanduser().resolve()
    OUT.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({'figure.dpi':150,'font.size':9,'axes.spines.top':False,'axes.spines.right':False})
    G=['Vehículo','SCOP-HB','LiCl'];REG=['BLAa','BLAp']
    a=pd.read_csv(R/'datos_animal.csv');p=pd.read_csv(R/'polaridad_animal.csv');glob=pd.read_csv(R/'densidad_principal.csv');dire=pd.read_csv(R/'direccion_2D.csv');cent=pd.read_csv(R/'centroides_2D.csv');hot=pd.read_csv(R/'maximos_3D.csv');cov=pd.read_csv(R/'cobertura_AP.csv')
    # Density by animal
    fig,axs=plt.subplots(1,2,figsize=(9,3.5),constrained_layout=True)
    for ax,reg in zip(axs,REG):
        for j,g in enumerate(G):
            v=a[(a.Region==reg)&(a.Grupo==g)].Densidad_celulas_mm3.to_numpy()/1000
            xs=j+np.linspace(-.12,.12,len(v));ax.scatter(xs,v,s=28,alpha=.8);ax.errorbar(j+.2,v.mean(),yerr=v.std(ddof=1),fmt='s',capsize=3)
        ax.set_title(reg);ax.set_xticks(range(3),[f'{g}\nn={sum((a.Region==reg)&(a.Grupo==g))}' for g in G]);ax.set_ylabel('Densidad (miles de objetos/mm³)');ax.grid(axis='y',alpha=.18)
    fig.savefig(OUT/'01_densidad_animales.png',bbox_inches='tight');plt.close(fig)
    # Descriptive polarity
    fig,axs=plt.subplots(2,2,figsize=(9,6),constrained_layout=True)
    for i,reg in enumerate(REG):
        for k,(num,den,label) in enumerate([('Lateral','Medial','Lateral / medial'),('Dorsal','Ventral','Dorsal / ventral')]):
            ax=axs[i,k]
            for j,g in enumerate(G):
                d=p[(p.Region==reg)&(p.Grupo==g)];v=d['Densidad_'+num]/d['Densidad_'+den]
                ax.scatter(j+np.linspace(-.12,.12,len(v)),v,s=24,alpha=.8);ax.plot([j-.16,j+.16],[v.median(),v.median()],lw=2)
            ax.axhline(1,ls='--',lw=.8);ax.set_yscale('log');ax.set_xticks(range(3),G);ax.set_title(f'{reg} · {label}');ax.set_ylabel('Razón de densidades');ax.grid(axis='y',alpha=.15)
    fig.savefig(OUT/'02_polaridad_2D_animales.png',bbox_inches='tight');plt.close(fig)
    # Direction forest
    fig,axs=plt.subplots(1,2,figsize=(9,4.2),constrained_layout=True)
    for ax,eje,title in zip(axs,['ML','DV'],['Cambio lateral/medial','Cambio dorsal/ventral']):
        d=dire[dire.Eje==eje].reset_index(drop=True);ys=np.arange(len(d))[::-1]
        for y,(_,r) in zip(ys,d.iterrows()):
            ax.errorbar(r.Efecto,y,xerr=[[r.Efecto-r.Efecto_IC95_inf],[r.Efecto_IC95_sup-r.Efecto]],fmt='o',capsize=3);ax.text(max(r.Efecto_IC95_sup,1)*1.05,y,f'pHolm={r.p_Holm:.3f}',va='center',fontsize=8)
        ax.axvline(1,ls='--',lw=.8);ax.set_xscale('log');ax.set_yticks(ys,[f'{r.Region} · {r.Tratamiento}' for _,r in d.iterrows()]);ax.set_title(title);ax.set_xlabel('Razón del balance vs vehículo')
    fig.savefig(OUT/'03_direccion_2D_ajustada.png',bbox_inches='tight');plt.close(fig)
    # Centroid forest
    fig,axs=plt.subplots(1,2,figsize=(9,4.2),constrained_layout=True)
    for ax,eje,title in zip(axs,['ML','DV'],['Centroide medial–lateral','Centroide ventral–dorsal']):
        d=cent[cent.Eje==eje].reset_index(drop=True);ys=np.arange(len(d))[::-1]
        for y,(_,r) in zip(ys,d.iterrows()):
            val=r.Efecto*100;lo=r.Efecto_IC95_inf*100;hi=r.Efecto_IC95_sup*100
            ax.errorbar(val,y,xerr=[[val-lo],[hi-val]],fmt='o',capsize=3);ax.text(max(hi,1)+1,y,f'pHolm={r.p_Holm:.3f}',va='center',fontsize=8)
        ax.axvline(0,ls='--',lw=.8);ax.set_yticks(ys,[f'{r.Region} · {r.Tratamiento}' for _,r in d.iterrows()]);ax.set_title(title);ax.set_xlabel('Cambio del eje normalizado (puntos %)')
    fig.savefig(OUT/'04_centroides_2D.png',bbox_inches='tight');plt.close(fig)
    # Equal-animal 2D maps
    x=XLSX(ROOT/'resultados_cfos_2D_3D.xlsx');b=x.read('Bins_2D_stack');b['Animal_ID']=b.Experimento.astype(str)+'::'+b.Animal.astype(str);b.Celulas=pd.to_numeric(b.Celulas,errors='coerce').fillna(0);b.Bin_X=pd.to_numeric(b.Bin_X,errors='coerce');b.Bin_Y=pd.to_numeric(b.Bin_Y,errors='coerce')
    for reg in REG:
        fig,axs=plt.subplots(1,3,figsize=(9,3.1),constrained_layout=True);maps=[]
        for g in G:
            mats=[]
            for aid,z in b[(b.Region==reg)&(b.Grupo==g)].groupby('Animal_ID'):
                m=np.zeros((30,30));
                for _,rr in z.groupby(['Bin_X','Bin_Y']).Celulas.sum().reset_index().iterrows():m[int(rr.Bin_Y),int(rr.Bin_X)]=rr.Celulas
                m=gaussian_filter(m,1.5,mode='reflect');m/=m.sum() if m.sum() else 1;mats.append(m)
            maps.append(np.mean(mats,axis=0))
        vmax=max(m.max() for m in maps)
        for ax,g,m in zip(axs,G,maps):
            im=ax.imshow(m,origin='lower',extent=[0,1,0,1],vmin=0,vmax=vmax,cmap='viridis');ax.axvline(.5,lw=.5);ax.axhline(.5,lw=.5);ax.set_title(g);ax.set_xlabel('Medial ← X → Lateral');ax.set_ylabel('Ventral ← Y → Dorsal')
        fig.colorbar(im,ax=axs,label='Probabilidad media por bin',shrink=.8);fig.savefig(OUT/f'05_mapa_2D_{reg}.png',bbox_inches='tight');plt.close(fig)
    # AP coverage
    for reg in REG:
        d=cov[cov.Region==reg];aps=sorted(d.Bregma_mm.unique(),reverse=True);M=np.zeros((3,len(aps)))
        for i,g in enumerate(G):
            z=d[d.Grupo==g].set_index('Bregma_mm').N_animales
            for j,ap in enumerate(aps):M[i,j]=z.get(ap,0)
        fig,ax=plt.subplots(figsize=(8,2.7));im=ax.imshow(M,aspect='auto',cmap='Blues');
        for i in range(3):
            for j in range(len(aps)):ax.text(j,i,str(int(M[i,j])),ha='center',va='center',fontsize=8)
        ax.set_yticks(range(3),G);ax.set_xticks(range(len(aps)),[f'{x:.2f}' for x in aps],rotation=45);ax.set_title(f'{reg}: animales con stack por bregma');ax.set_xlabel('Bregma (mm)');fig.colorbar(im,ax=ax,label='Animales');fig.tight_layout();fig.savefig(OUT/f'06_cobertura_AP_{reg}.png',bbox_inches='tight');plt.close(fig)
    # 3D hotspot support
    fig,ax=plt.subplots(figsize=(8,4));h=hot.copy();labels=[f'{r.Region} · {r.Grupo}' for _,r in h.iterrows()];y=np.arange(len(h));ax.barh(y-.15,h.Animales_con_volumen,height=.28,label='Animales con volumen');ax.barh(y+.15,h.Animales_con_celulas,height=.28,label='Animales con objetos');ax.set_yticks(y,labels);ax.set_xlabel('Número de animales en el voxel máximo');ax.legend();ax.invert_yaxis();fig.tight_layout();fig.savefig(OUT/'07_soporte_maximos_3D.png',bbox_inches='tight');plt.close(fig)
    print('figures',len(list(OUT.glob('*.png'))))


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True, type=Path, help='Carpeta de la corrida original')
    parser.add_argument('--results', required=True, type=Path, help='Resultados estadísticos existentes')
    parser.add_argument('--output', required=True, type=Path, help='Carpeta de figuras')
    args=parser.parse_args()
    run(args.input, args.results, args.output)


if __name__ == "__main__":
    main()

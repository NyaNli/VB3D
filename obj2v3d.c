#include<stdio.h>
#include<string.h>
#include<stdlib.h>

#pragma pack(push)
#pragma pack(1)
typedef struct
{
	double x,y,z;
	short r,g,b;
} point;
#pragma pack(pop)

typedef struct
{
	point p[3];
} delta;

void skipline(FILE *fp)
{
	int c;
	if(feof(fp))
		return;
	do
	{
		c = fgetc(fp);
	}while(c!='\n' && !feof(fp));
}

int main(int argv, char**args)
{
	char c;
	int vs=0,ds=0,i;
	int p1,p2,p3;
	int rr,gg,bb;
	FILE *fp;
	point *p;
	delta *d;
	if(argv!=2)
		return 0;
	fp = fopen(args[1],"r");
	while(!feof(fp))
	{
		c = fgetc(fp);
		if(c=='v')
			vs++;
		else if(c=='f')
			ds++;
		else if(c!='\n')
			skipline(fp);
	}
	printf("%d %d\n",vs,ds);
	p = (point*)malloc(vs*sizeof(point));
	d = (delta*)malloc(ds*sizeof(delta));
	fseek(fp,0,0);
	i = 0;
	rr=255;
	gg=255;
	bb=255;
	while(!feof(fp))
	{
		c = fgetc(fp);
		// Custom Label for Vertex Color
		if(c=='c')
			fscanf(fp,"%d%d%d\n",&rr,&gg,&bb);
		else if(c=='v')
		{
			fscanf(fp,"%lf%lf%lf\n",&p[i].x,&p[i].y,&p[i].z);
			p[i].r = rr;
			p[i].g = gg;
			p[i].b = bb;
			i++;
		}
		else if(c!='\n')
			skipline(fp);
	}
	fseek(fp,0,0);
	i = 0;
	while(!feof(fp))
	{
		c = fgetc(fp);
		if(c=='f')
		{
			fscanf(fp,"%d%d%d\n",&p1,&p2,&p3);
			d[i].p[0] = p[p1-1];
			d[i].p[1] = p[p2-1];
			d[i].p[2] = p[p3-1];
			i++;
		}
		else if(c!='\n')
			skipline(fp);
	}
	fclose(fp);
	i = strlen(args[1]);
	args[1][i-3]='v';
	args[1][i-2]='3';
	args[1][i-1]='d';
	fp = fopen(args[1],"wb");
	/*
	for(i=0;i<ds;i++)
	{
		for(j=0;j<3;j++)
			printf("X:%.5f  Y:%.5f  Z:%.5f\nR:%d G:%d B:%d\n",d[i].p[j].x,d[i].p[j].y,d[i].p[j].z,d[i].p[j].r,d[i].p[j].g,d[i].p[j].b);
		putchar('\n');
	}
	*/
	fwrite(d,sizeof(delta),ds,fp);
	fclose(fp);
	return 0;
}

/*
int main()
{
	int i,j;
	FILE *fp = fopen("E:\\Desktop\\box.v3d","rb");
	delta d[12];
	fread(d,sizeof(delta),12,fp);
	for(i=0;i<12;i++)
	{
		for(j=0;j<3;j++)
			printf("X:%.5f  Y:%.5f  Z:%.5f\nR:%d G:%d B:%d\n",d[i].p[j].x,d[i].p[j].y,d[i].p[j].z,d[i].p[j].r,d[i].p[j].g,d[i].p[j].b);
		putchar('\n');
	}
	fclose(fp);
	return 0;
}*/
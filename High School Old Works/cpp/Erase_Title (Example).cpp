/*----- ----- ----- ----- ----- -----*/
//      Erasing Title (Example)      //
//                                   //
//Auther : Synasaivaltos Terdolic    //
//Company:                           //
//Date   : 2014/12/31                //
//                                   //
//      synasaivaltos@gmail.com      //
/*----- ----- ----- ----- ----- -----*/

#include <conio.h>
#include <cstdlib>
#include <iostream>
#include <windows.h>

#define erased_char ' '
#define erasing_char1 1
#define erasing_char2 ' '
#define time 100
#define title_col 71
#define title_line 7

using namespace std;

void pause(void);
void initialize_basic_frame(void);
void basic_frame(void);
void erase_title(void);
void refresh_title(void);
void show_basic_frame(void);

char title_char[title_line][title_col];

int main(void)
{
	extern char title_char[title_line][title_col];
	
	while(1)
	{
		initialize_basic_frame();
		basic_frame();
		show_basic_frame();
		Sleep(5000);
		erase_title();
		pause();
		system("cls");
	}
	
	pause();
	return 0;
}

void pause(void)
{
	cout << "Press any key to continue...";
	char temp=getch();
	
	return;
}

void initialize_basic_frame(void)
{
	//initialize "title_char"
	for(int i=0;i<title_line;i++)
	{
		for(int j=0;j<title_col;j++)
		{
			title_char[i][j]=' ';
		}
	}
	
	return;
}

void basic_frame(void)
{
	for(int i=0;i<title_col;i++)
	{
		if((i+1)%6==0)
		{
			title_char[0][i]=' ';
			title_char[2][i]=' ';
			title_char[6][i]=' ';
		}
		else
		{
			title_char[0][i]='-';
			title_char[2][i]='-';
			title_char[6][i]='-';
		}
	}
	title_char[1][0]='|';
	title_char[1][24]='E';
	title_char[1][25]='r';
	title_char[1][26]='a';
	title_char[1][27]='s';
	title_char[1][28]='i';
	title_char[1][29]='n';
	title_char[1][30]='g';
	title_char[1][32]='T';
	title_char[1][33]='i';
	title_char[1][34]='t';
	title_char[1][35]='l';
	title_char[1][36]='e';
	title_char[1][38]='(';
	title_char[1][39]='E';
	title_char[1][40]='x';
	title_char[1][41]='a';
	title_char[1][42]='m';
	title_char[1][43]='p';
	title_char[1][44]='l';
	title_char[1][45]='e';
	title_char[1][46]=')';
	title_char[1][70]='|';
	title_char[3][0]='|';
	title_char[3][21]='M';
	title_char[3][22]='a';
	title_char[3][23]='d';
	title_char[3][24]='e';
	title_char[3][26]='b';
	title_char[3][27]='y';
	title_char[3][29]='S';
	title_char[3][30]='y';
	title_char[3][31]='n';
	title_char[3][32]='a';
	title_char[3][33]='s';
	title_char[3][34]='a';
	title_char[3][35]='i';
	title_char[3][36]='v';
	title_char[3][37]='a';
	title_char[3][38]='l';
	title_char[3][39]='t';
	title_char[3][40]='o';
	title_char[3][41]='s';
	title_char[3][43]='T';
	title_char[3][44]='e';
	title_char[3][45]='r';
	title_char[3][46]='d';
	title_char[3][47]='o';
	title_char[3][48]='l';
	title_char[3][49]='i';
	title_char[3][50]='c';
	title_char[3][70]='|';
	title_char[4][0]='|';
	title_char[4][24]='s';
	title_char[4][25]='y';
	title_char[4][26]='n';
	title_char[4][27]='a';
	title_char[4][28]='s';
	title_char[4][29]='a';
	title_char[4][30]='i';
	title_char[4][31]='v';
	title_char[4][32]='a';
	title_char[4][33]='l';
	title_char[4][34]='t';
	title_char[4][35]='o';
	title_char[4][36]='s';
	title_char[4][37]='@';
	title_char[4][38]='g';
	title_char[4][39]='m';
	title_char[4][40]='a';
	title_char[4][41]='i';
	title_char[4][42]='l';
	title_char[4][43]='.';
	title_char[4][44]='c';
	title_char[4][45]='o';
	title_char[4][46]='m';
	title_char[4][70]='|';
	title_char[5][0]='|';
	title_char[5][32]='V';
	title_char[5][33]='e';
	title_char[5][34]='r';
	title_char[5][35]='.';
	title_char[5][37]='1';
	title_char[5][38]='.';
	title_char[5][39]='0';
	title_char[5][70]='|';
	
	return;
}

void erase_title(void)
{
	for(int i=0;i<title_line;i++)
	{
		if(title_col%2==1)  //
		{
			for(int j=0;j<title_col-1;j+=2)
			{
				if(j==0 && i!=0)
				{
					title_char[i-1][title_col-1]=erased_char;
				}
				if(j>0)
				{
					title_char[i][j-2]=erased_char;
					title_char[i][j-1]=erased_char;
				}
				title_char[i][j]=erasing_char1;
				title_char[i][j+1]=erasing_char2;
				refresh_title();
			}
			title_char[i][title_col-3]=erased_char;
			title_char[i][title_col-2]=erased_char;
			title_char[i][title_col-1]=erasing_char1;
			refresh_title();
		}
		else  //
		{
			for(int j=0;j<title_col;j+=2)
			{
				if(j==0 && i!=0)
				{
					title_char[i-1][title_col-2]=erased_char;
				}
				if(j>0)
				{
					title_char[i][j-2]=erased_char;
					title_char[i][j-1]=erased_char;
				}
				title_char[i][j]=erasing_char1;
				title_char[i][j+1]=erasing_char2;
				refresh_title();
			}
		}
	}
	title_char[title_line-1][title_col-2]=erased_char;
	title_char[title_line-1][title_col-1]=erased_char;
	refresh_title();
	return;
}

void refresh_title(void)
{
	system("cls");
	show_basic_frame();
	Sleep(time);
	
	return;
}
void show_basic_frame(void)
{
	for(int i=0;i<title_line;i++)
	{
		for(int j=0;j<title_col;j++)
		{
			cout << title_char[i][j];
		}
		cout << endl;
	}
	
	return;
}

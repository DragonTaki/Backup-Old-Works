/*----- ----- ----- ----- ----- -----*/
//           Game Of Life            //
//                                   //
//Auther : Synasaivaltos Terdolic    //
//Company:                           //
//Date   : 2014/12/30                //
/*----- ----- ----- ----- ----- -----*/

#include <conio.h>
#include <cstdlib>
#include <ctime>
#include <iostream>
#include <windows.h>

#define cell_length 10
#define cell_width 26
#define ALIVE 1
#define DEAD 0

using namespace std;

void pause(void);
void basic_frame(void);
void rule_frame(void);
void game_frame(void);
void manual_setting_mode(void);
void random_mode(void);
void game_start(void);
void print_cells(void);

bool cells[cell_length][cell_width];
bool first_execute=true;

int main(void)
{
	extern bool cells[cell_length][cell_width];
	extern bool first_execute;
	char temp=' ';
	
	while(1)
	{
		//initialize "temp"
		temp=' ';
		
		if(first_execute==true)
		{
			basic_frame();
			rule_frame();
			while(temp!='0' && temp!='1')
			{
				temp=getch();
			}
			if(temp=='0')  //manual_setting_mode
			{
				first_execute=false;
				basic_frame();
				game_frame();
				manual_setting_mode();
			}
			else  //random_mode
			{
				random_mode();
			}
		}
		
	}
	
	system("pause");
	return 0;
}

void pause(void)
{
	cout << "Press any key to continue...";
	char temp=getch();
	
	return;
}

void basic_frame(void)
{
	system("cls");
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << "|                            Game Of Life                             |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << "|                    Made by Synasaivaltos Terdolic                   |" << endl;
	cout << "|                       synasaivaltos@gmail.com                       |" << endl;
	cout << "|                               Ver. 1.1                              |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	
	return;
}

void rule_frame(void)
{
	cout << "|*The universe of the Game of Life is an infinite two-dimensional     |" << endl;
	cout << "|orthogonal grid of square cells, each of which is in one of two      |" << endl;
	cout << "|possible states, alive or dead. Every cell interacts with its eight  |" << endl;
	cout << "|neighbours, which are the cells that are horizontally, vertically, or|" << endl;
	cout << "|diagonally adjacent. At each step in time, the following transitions |" << endl;
	cout << "|occur:                                                               |" << endl;
	cout << "|1. Any live cell with fewer than two live neighbours dies, as if     |" << endl;
	cout << "|   caused by under-population.                                       |" << endl;
	cout << "|2. Any live cell with two or three live neighbours lives on to the   |" << endl;
	cout << "|   next generation.                                                  |" << endl;
	cout << "|3. Any live cell with more than three live neighbours dies, as if by |" << endl;
	cout << "|   overcrowding.                                                     |" << endl;
	cout << "|4. Any dead cell with exactly three live neighbours becomes a live   |" << endl;
	cout << "|   cell, as if by reproduction.                                      |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << "  0  Manual Setting Mode              1  Random Mode                   " << endl;
	
	return;
}

void game_frame(void)
{
	cout << "|         A B C D E F G H I J K L M N O P Q R S T U V W X Y Z         |" << endl;
	cout << "|        ------------------------------------------------------       |" << endl;
	for(int i=0;i<cell_length;i++)
	{
		cout << "|      " << i << " |";
		for(int j=0;j<cell_width;j++)
		{
			if(cells[i][j]==DEAD)
			{
				cout << "¡¼";
			}
			else  //ALIVE
			{
				cout << "¡½";
			}
		}
		cout << "| " << i << "     |" << endl;
	}
	cout << "|        ------------------------------------------------------       |" << endl;
	cout << "|         A B C D E F G H I J K L M N O P Q R S T U V W X Y Z         |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	
	return;
}

void manual_setting_mode(void)
{
	char width=' ';  //x-axis
	char length=' ';  //y-axis
	char temp=' ';
	
	//initialize "cells"
	for(int i=0;i<cell_length;i++)
	{
		for(int j=0;j<cell_width;j++)
		{
			cells[i][j]=DEAD;
		}
	}
	//print_cells(cells);
	//pause();
	
	//get change status coordinates
	while(1)
	{
		//initialize "width", "length", "temp"
		width=' ';
		length=' ';
		temp=' ';
		//start to change status
		cout << "Input x-axis (A~Z): ";
		while(width<'A' || width>'Z')
		{
			temp=getch();
			if(temp==13 || temp==27)
			{
				cout << "                ";
				break;
			}
			if(temp>'a'-1 && temp<'z'+1)
			{
				temp-=32;
			}
			if(temp>'A'-1 && temp<'Z'+1)
			{
				cout << temp << "               ";
				width=temp;
			}
		}
		cout << "Input y-axis (0~9): ";
		while(length<'0' || length>'9')
		{
			if(temp==13 || temp==27)
			{
				cout << endl;
				break;
			}
			temp=getch();
			if(temp>'0'-1 && temp<'9'+1)
			{
				cout << temp << endl;
				length=temp;
			}
		}
		if(temp!=13 && temp!=27)
		{
			if(cells[length-48][width-65]==ALIVE)
			{
				cells[length-48][width-65]=DEAD;
			}
			else
			{
				cells[length-48][width-65]=ALIVE;
			}
			//print_cells(cells);
			pause();
			basic_frame();
			game_frame();
		}
		else if(temp==13)  //end input coordinates
		{
			cout << "End change cell status? (Y/N)";
			while(temp!='Y' && temp!='y' && temp!='N' && temp!='n')
			{
				temp=getch();
				if(temp=='Y' || temp=='y')
				{
					break;
				}
				else if(temp=='N' || temp=='n')
				{
					temp=' ';
					basic_frame();
					game_frame();
					break;
				}
			}
		}
		else if(temp==27)  //press ESC
		{
			first_execute=true;
			break;
		}
		if(temp=='Y' || temp=='y')
		{
			//start the game
			game_start();
			break;
		}
	}
	
	return;
}

void random_mode(void)
{
	srand(int(time(0)));
	
	int temp;
	
	for(int i=0;i<cell_length;i++)
	{
		for(int j=0;j<cell_width;j++)
		{
			temp=rand()%8;
			if(temp==2 || temp==3)
			{
				cells[i][j]=ALIVE;
			}
			else
			{
				cells[i][j]=DEAD;
			}
		}
	}
	game_start();
	
	return;
}

void game_start(void)
{
	bool temp_cells[cell_length][cell_width];
	int around_cells_ALIVE_num=0;
	int i_up,i_down;
	int j_left,j_right;
	bool color_default=true;
	
	while(1)
	{
		Sleep(1000);  //status keep 1 second
		//check adjacent
		for(int i=0;i<cell_length;i++)
		{
			for(int j=0;j<cell_width;j++)
			{
				//check up and down
				if(i==0)  //top
				{
					i_up=cell_length-1;
					i_down=i+1;
				}
				else if(i==cell_length-1)  //bottom
				{
					i_up=i-1;
					i_down=0;
				}
				else
				{
					i_up=i-1;
					i_down=i+1;
				}
				//check left and right
				if(j==0)  //leftmost
				{
					j_left=cell_width-1;
					j_right=j+1;
				}
				else if(j==cell_width-1)  //rightmost
				{
					j_left=j-1;
					j_right=0;
				}
				else
				{
					j_left=j-1;
					j_right=j+1;
				}
				//print_cells(cells);
				//pause();
				
				//check adjacent status
				if(cells[i_up][j_left]==ALIVE)  //left up
				{
					around_cells_ALIVE_num++;
				}
				if(cells[i_up][j]==ALIVE)  //up
				{
					around_cells_ALIVE_num++;
				}
				if(cells[i_up][j_right]==ALIVE)  //right up
				{
					around_cells_ALIVE_num++;
				}
				if(cells[i][j_left]==ALIVE)  //left
				{
					around_cells_ALIVE_num++;
				}
				if(cells[i][j_right]==ALIVE)  //right
				{
					around_cells_ALIVE_num++;
				}
				if(cells[i_down][j_left]==ALIVE)  //left down
				{
					around_cells_ALIVE_num++;
				}
				if(cells[i_down][j]==ALIVE)  //down
				{
					around_cells_ALIVE_num++;
				}
				if(cells[i_down][j_right]==ALIVE)  //right down
				{
					around_cells_ALIVE_num++;
				}
				/*//debug
				cout << "cells[i_up][j_left] = " << cells[i_up][j_left] << endl;
				cout << "cells[i_up][j] = " << cells[i_up][j] << endl;
				cout << "cells[i_up][j_right] = " << cells[i_up][j_right] << endl;
				cout << "cells[i][j_left] = " << cells[i][j_left] << endl;
				cout << "cells[i][j_right] = " << cells[i][j_right] << endl;
				cout << "cells[i_down][j_left] = " << cells[i_down][j_left] << endl;
				cout << "cells[i_down][j] = " << cells[i_down][j] << endl;
				cout << "cells[i_down][j_right] = " << cells[i_down][j_right] << endl;
				cout << "around_cells_ALIVE_num = " << around_cells_ALIVE_num << endl;*/
				//judge cell will ALIVE or DEAD
				if(around_cells_ALIVE_num==3)  //cell will ALIVE
				{
					temp_cells[i][j]=ALIVE;
				}
				else if(around_cells_ALIVE_num==2)  //keep status
				{
					temp_cells[i][j]=cells[i][j];
				}
				else  //cell will DEAD
				{
					temp_cells[i][j]=DEAD;
				}
				around_cells_ALIVE_num=0;
			}
		}
		//next ststus
		for(int i=0;i<cell_length;i++)
		{
			for(int j=0;j<cell_width;j++)
			{
				cells[i][j]=temp_cells[i][j];
			}
		}
		basic_frame();
		game_frame();
		if(color_default==true)
		{
			system("color F");
			color_default=false;
		}
		else
		{
			system("color 7");
			color_default=true;
		}
	}
	
	return;
}

void print_cells(void)
{
	for(int i=0;i<cell_length;i++)
	{
		for(int j=0;j<cell_width;j++)
		{
			if(cells[i][j]==DEAD)
			{
				cout << "0";
			}
			else
			{
				cout << "1";
			}
		}
		cout << endl;
	}
	
	return;
}

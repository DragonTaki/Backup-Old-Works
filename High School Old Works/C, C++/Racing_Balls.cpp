/*----- ----- ----- ----- ----- -----*/
//            Racing Balls           //
//                                   //
//Auther : Synasaivaltos Terdolic    //
//Company:                           //
//Date   : 2014/12/27                //
/*----- ----- ----- ----- ----- -----*/

#include <cmath>
#include <conio.h>
#include <cstdlib>
#include <iomanip>
#include <iostream>

#define pi 3.1415926 
#define g 9.8

using namespace std;

void basic_frame(void);

int main(void)
{
	float v,d,a,l,k;
	float t1,t2;
	
	basic_frame();
	
	while(1)
	{
		v=0;
		while(v<=0)
		{
			cout << "Initial velocity (m/s): ";
			cin >> v;
			if(v<=0)
	      {
	      	system("cls");
	      	basic_frame();
	  			cout << "Initial velocity must be greater than zero!" << endl;
			}
		}

		d=0;
		while(d<=0)
		{
			cout << "The length of the slope (m): ";
			cin >> d;
			if(d<=0)
	      {
	      	system("cls");
	      	basic_frame();
				cout << "Initial velocity (m/s): " << v << endl << endl;
	  			cout << "The length of the slope must be greater than zero!" << endl;
			}
		}
		
		l=0;
 	   while(l<=0)
 	   {
			cout << "The level length after roll the slope (m): ";
			cin >> l;
			if(l<=0)
	      {
	      	system("cls");
	      	basic_frame();
				cout << "Initial velocity (m/s): " << v << endl;
				cout << "The length of the slope (m): " << d << endl << endl;
	  			cout << "The level length must be greater than zero!" << endl;
			}
 	   }
 	   
 	   a=0;
 	   while(a<=0 or a>=90)
 	   {
			cout << "The angle of the slope (deg): ";
			cin >> a;
			if(a<=0 or a>=90)
	      {
	      	system("cls");
	      	basic_frame();
				cout << "Initial velocity (m/s): " << v << endl;
				cout << "The length of the slope (m): " << d << endl;
				cout << "The level length after roll the slope (m): " << l << endl << endl;
	  			cout << "The angle must be between 0~90 degrees!" << endl;
			}
 	   }
		
		k=sqrt(v*v+g*d*tan(a*pi/180));
		t1=(2*(k-v)/(g*sin(a*pi/180)))+(l/k);
		t2=((l+d)/v);
		
   	system("cls");
   	basic_frame();
		cout << "Initial velocity (m/s): " << v << endl;
		cout << "The length of the slope (m): " << d << endl;
		cout << "The level length after roll the slope (m): " << l << endl;
		cout << "The angle of the slope (deg): " << a << endl << endl;
		cout << "The time pass the slope (s): " << t1 << endl;
		cout << "The time pass the ground (s): " << t2 << endl << endl;
		
		cout << "0 for exit, 1 for continue." << endl;
		char temp;
		temp=getch();
		if(temp=='0')
		{
			return 0;
		}
		else
		{
	   	system("cls");
	   	basic_frame();
		}
	}
	
	system("pause");
	return 0;
}

void basic_frame(void)
{
	system("cls");
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << "|                             Racing Balls                            |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << "|                    Made by Synasaivaltos Terdolic                   |" << endl;
	cout << "|                       synasaivaltos@gmail.com                       |" << endl;
	cout << "|                               Ver. 1.0                              |" << endl;
	cout << "----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- -----" << endl;
	cout << endl;
	
	return;
}

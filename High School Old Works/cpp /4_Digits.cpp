/*----- ----- ----- ----- ----- -----*/
//              4 Digits             //
//                                   //
//Auther : Synasaivaltos Terdolic    //
//Company:                           //
//Date   : 2014/12/27                //
/*----- ----- ----- ----- ----- -----*/

#include <cstdlib>
#include <conio.h>
#include <ctime>
#include <iostream>

#define ans_digits 4
#define ans_limit 10
#define avl_input_char_num 1000
#define BAC_num_sum 10

using namespace std;

void frame(void);
void show_ans(void);
void char_to_num(char input,int ans_count,int i);
void nAnB(int ans[ans_limit][ans_digits],int ans_count);

int main(void)
{
	extern int BAC_ans[BAC_num_sum];  //the forst four number are answers
	extern int ans[ans_limit][ans_digits];
	extern int ans_count;
	extern int A[ans_limit];
	extern int B[ans_limit];
	char input[avl_input_char_num];
	bool valid_ans=true;
	bool only_input_4_char=true;
	bool correct_ans=false;
	bool showed_ans=false;
	
	srand(int(time(0)));
	
	while(1)
	{
		//initialize "ans_count", "showed_ans", "A", "B", "ans"
		ans_count=0;
		showed_ans=false;
		for(int i=0;i<ans_limit;i++)
		{
			A[i]=0;
			B[i]=0;
		}
		for(int i=0;i<ans_limit;i++)
		{
			for(int j=0;j<ans_digits;j++)
			{
				ans[i][j]=0;
			}
		}
		
		frame();
		
		for(int i=0;i<BAC_num_sum;i++)  //make BAC number
		{
			BAC_ans[i]=i;
		}
		for(int i=9;i>0;i--)
		{
			int pos;
			int temp;
			pos=rand()%10;
			temp=BAC_ans[i];
			BAC_ans[i]=BAC_ans[pos];
			BAC_ans[pos]=temp;
		}
		/*for(int i=0;i<10;i++)
		{
			cout << "Num[" << i+1 << "] = " << BAC_ans[i] << endl;
		}*/ 
		
		while(ans_count<ans_limit)  //input answers
		{
			//initialize "valid_ans", "input"
			valid_ans=true;
			for(int i=0;i<avl_input_char_num;i++)
			{
				input[i]='\0';
				//cout << input[i];
			}
			//get an answer
			cout << "----- ----- -----" << endl;
			cout << "Ans:";
			cin.getline(input,avl_input_char_num);
			//check the answer
			if(input[3]!='\0' && input[4]=='\0')  //only input four char
			{
				for(int i=0;i<ans_digits;i++)
				{
					if(input[i]<48 || input[i]>57)  //not a number
					{
						valid_ans=false;
						break;
					}
				}
				for(int i=0;i<ans_digits;i++)
				{
					for(int j=i;j<ans_digits;j++)
					{
						if(i!=j && input[i]==input[j])  //number duplicate
						{
							valid_ans=false;
							break;
						}
					}
				}
			}
			else //invalid char length
			{
				if(input[0]=='a' && input[1]=='n' && input[2]=='s')
				{
					show_ans();
					showed_ans=true;
				}
				valid_ans=false;
			}
			if(valid_ans==true)  //write input to answer array
			{
				for(int i=0;i<ans_digits;i++)
				{
					char_to_num(input[i],ans_count,i);
				}
				nAnB(ans,ans_count);  //calculate A, B
				ans_count++;
				frame();
				
				if(A[ans_count-1]==ans_digits)  //correct answer
				{
					correct_ans=true;
					break;
				}
			}
			else if(showed_ans==false)  //not an valid answer
			{
				frame();
				cout << "*Not a valid answer!" << endl << "*Please input your answer again!" << endl;
			}
		}
		//determine if answered
		if(correct_ans==true)
		{
			cout << "Well done!" << endl << "You have correct answer!" << endl;
		}
		else
		{
			show_ans();
			showed_ans=true;
			cout << "Sorry, you had not got the correct answer!" << endl;
		}
		cout << endl << endl << "Play again? (Y/N)" << endl;
		char temp=' ';
		while(temp!='Y' && temp!='y')
		{
			temp=getch();
			if(temp=='Y' || temp=='y')
			{
				system("color 7");
			}
			else if(temp=='N' || temp=='n')
			{
				return 0;
			}
		}
	}
	
	system("pause");
	return 0;
}

int BAC_ans[BAC_num_sum];
int ans[ans_limit][ans_digits];
int ans_count=0;
int A[ans_limit];
int B[ans_limit];

void frame(void)
{
	system("cls");
	cout << "----- ----- -----" << endl;
	cout << "4 Digits" << endl << endl;
	
	for(int i=0;i<ans_count;i++)
	{
		if(int((i+1)/10.0)==0)
		{
			cout << " ";
		}
		cout << i+1 << ":  ";
		for(int j=0;j<ans_digits;j++)
		{
			cout << ans[i][j];
		}
		cout << "    " << A[i] << "A" << B[i] << "B" << endl;
	}
	
	for(int i=ans_count;i<ans_limit;i++)
	{
		cout << endl;
	}
	cout << endl << endl;
	
	return;
}

void show_ans(void)
{
	system("cls");
	system("color C");
	cout << "----- ----- -----" << endl;
	cout << "4 Digits" << endl << endl;
	
	for(int i=0;i<ans_limit;i++)
	{
		if(int((i+1)/10.0)==0)
		{
			cout << " ";
		}
		cout << i+1 << ":  ";
		if(i<ans_count)
		{
			for(int j=0;j<ans_digits;j++)
			{
				cout << ans[i][j];
			}
			cout << "    " << A[i] << "A" << B[i] << "B" << endl;
		}
		else
		{
			cout << endl;
		}
	}
	cout << endl << "Ans: ";
	for(int i=0;i<ans_digits;i++)
	{
		cout << BAC_ans[i];
	}
	cout << endl;
	ans_count=ans_limit;
	
	return;
}

void char_to_num(char input,int ans_count,int i)
{
	switch(input)
	{
		case '0':
			ans[ans_count][i]=0;
			break;
		case '1':
			ans[ans_count][i]=1;
			break;
		case '2':
			ans[ans_count][i]=2;
			break;
		case '3':
			ans[ans_count][i]=3;
			break;
		case '4':
			ans[ans_count][i]=4;
			break;
		case '5':
			ans[ans_count][i]=5;
			break;
		case '6':
			ans[ans_count][i]=6;
			break;
		case '7':
			ans[ans_count][i]=7;
			break;
		case '8':
			ans[ans_count][i]=8;
			break;
		case '9':
			ans[ans_count][i]=9;
			break;
	}
	
	return;
}

void nAnB(int ans[ans_limit][ans_digits],int ans_count)
{
	//calculate A
	for(int i=0;i<ans_digits;i++)
	{
		if(ans[ans_count][i]==BAC_ans[i])
		{
			A[ans_count]++;
		}
	}
	//calculate B
	for(int i=0;i<ans_digits;i++)
	{
		for(int j=0;j<ans_digits;j++)
		{
			if(ans[ans_count][i]==BAC_ans[j] && i!=j)
			{
				B[ans_count]++;
			}
		}
	}
	
	return;
}

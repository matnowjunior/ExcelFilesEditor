#include <string>
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <filesystem>


//define colors
#define RED "\033[31m"    
#define GREEN "\033[32m" 
#define RESET "\033[0m"

namespace fs = std::filesystem;

using namespace std;

//global variables
string announcement, color;

string wybierzArkusz(xlnt::workbook& wb)
{
    xlnt::worksheet ws;
    int num_sheets = wb.sheet_count();//returns number of workshhets in this workbook
    cout << "Avaiable sheets:" << endl;
    for (int i = 1; i <= num_sheets; i++) {
        ws = wb.sheet_by_index(i - 1);//returns worksheet at given index
        cout << i << ". " << ws.title() << endl;//returns title of the sheet
    }

    int wybor;
    while (true) {
        cout << "Choose sheet number: "<<endl;
        cin >> wybor;
        if (wybor >= 1 && wybor <= num_sheets) {
            xlnt::worksheet wybranyArkusz = wb.sheet_by_index(wybor - 1);
            return wybranyArkusz.title();
        }
        else {
            cout << "Sheet number is not valid. Try again." << endl;
        }
    }
}

string CopyFile()
{
    string originalFileName, NewFileName, UserName;

    do
    {
        cout << "Provide original path name: ";
        cin >> originalFileName;

        if (!fs::exists(originalFileName))
        {
            cout << RED << "Error occured" << RESET << endl;
        }
        else
        {
            cout << GREEN << "Correct path" << RESET << endl;
        }
    } while (!fs::exists(originalFileName));

    do
    {
        cout << "Provide new path: ";
        cin >> NewFileName;

        if (!fs::exists(NewFileName))
        {
            cout << RED << "Error occured" << RESET << endl;
        }
        else
        {
            cout << GREEN << "Correct path" << RESET << endl;
        }
    } while (!fs::exists(NewFileName));

    cout << "provide new file name: ";
    cin >> UserName;

    string DestName = NewFileName + "\\" + UserName;

    fs::copy_file(originalFileName, DestName);

    return DestName;

}

int titleToNumber(string s)
{
    int r = 0;
    for (int i = 0; i < s.length(); i++)
    {
        r = r * 26 + s[i] - 64;
    }
    return r;
}

pair <int, int> signs_numbers_separately(string cell_input)
{
    string letters, digits;
    for (char c : cell_input) {
        if (isalpha(c)) {
            letters += c;  //adding letter to letter variable
        }
        else if (isdigit(c)) {
            digits += c;  //adding number to digits variable
        }
    }

    return make_pair(titleToNumber(letters), stoi(digits));

}

int main()
{
    int i, j;
    string newFileName, originalFileName, cell_input1, cell_input2, newName;

    newName = CopyFile();



    //Opening Excel workbook and worksheet
    xlnt::workbook wb;
    xlnt::worksheet ws;

    //Loading excel file
    try {
        wb.load(newName);  // excel file path

        ws = wb.sheet_by_title(wybierzArkusz(wb));//getting specified by user worksheet from workbook

    }
    catch (const xlnt::exception& e) {
        cout << RED << "Processing failed " << RESET << e.what() << endl;//displaying an error message when trying to read a file
        return 1;
    }



    cout << "Podaj pierwsza komorke:";
    cin >> cell_input1;


    int col_number1 = signs_numbers_separately(cell_input1).first;
    int row_number1 = signs_numbers_separately(cell_input1).second;

    cout << "Podaj druga komorke:";
    cin >> cell_input2;

    int col_number2 = signs_numbers_separately(cell_input2).first;
    int row_number2 = signs_numbers_separately(cell_input2).second;


    //looping through (before) specified range
    for (i = row_number1; i < row_number2 + 1; i++)
    {

        for (j = col_number1; j < col_number2 + 1; j++)
        {
            //setting a new cell to the entered values(col and row)
            xlnt::cell cell = ws.cell(i, j);
            float x = cell.value<float>();


            if (cell.has_value())
            {
                cout << cell.value<int>();

                if (x >= 0 && x < 30)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(242, 34, 15)));
                else if (x >= 30 && x < 50)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(242, 160, 7)));
                else if (x >= 50 && x < 75)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(5, 242, 108)));
                else if (x >= 75 && x < 90)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(5, 151, 242)));
                else if (x >= 90 && x < 100)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(178, 37, 217)));
                cout << " ";
            }
            else
            {
                cout << "0";
                cout << " ";
            }

            //cell.clear_style();


        }
        cout << endl;
    }



    wb.save(newName);
    cout << announcement << endl;

    cout << GREEN << "Processing succeed :)" << RESET;
}
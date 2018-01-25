#include <iostream>
#include <fstream>
#include <sstream> 
#include <vector>
#include <iterator>
#include <algorithm>
#include <time.h> 

#include "xlsxwriter.h"

using namespace std;
using std::cout;
using std::endl;

int getFileSize(const std::string &fileName) {
	ifstream file(fileName.c_str(), ifstream::in | ifstream::binary);

	if (!file.is_open()) {
		return -1;
	}

	file.seekg(0, ios::end);
	int fileSize = file.tellg();
	file.close();

	return fileSize;
}

int main(int argc, char* argv[]) {

	//file
	const std::string csvFilePath =  argv[1];
	std::ifstream file(csvFilePath);
	if (!file) {
		return -1;
	}

	//excel
	const std::string excelFilePath = csvFilePath.substr(0, csvFilePath.find_last_of('.')) + ".xlsx";
	lxw_workbook  *workbook = new_workbook(excelFilePath.c_str());
	lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Sheet1");
	lxw_format *bold = workbook_add_format(workbook);
	format_set_bold(bold);

	//timer
	time_t start, end;
	time(&start);

	//import
	std::size_t row = 0;
	std::string line;
	while (std::getline(file, line)) {

		if (line.length() == 0)
			continue;

		std::string value;
		std::istringstream stream(line);
		std::size_t column = 0;

		while (std::getline(stream, value, ';')) {

			if (value.size() > 0) {

				if (value.find_first_of("'") == 0)
					value.erase(0, 1);

				if (value.find_last_of("'") == value.size() - 1)
					value.erase(value.size() - 1, 1);
			}

			worksheet_write_string(worksheet, row, column, value.c_str(), (row == 0 ? bold : NULL));

			column++;
		}

		row++;
	}

	//clean
	file.close();
	workbook_close(workbook);
	time(&end);

	//log
	std::ofstream log;
	const std::string logFilePath = csvFilePath.substr(0, csvFilePath.find_last_of('.')) + ".log";
	log.open(logFilePath);
	log << "Input file: " << csvFilePath << endl;
	log << "Output size: " << getFileSize(excelFilePath) << " bytes" << endl;
	log << "Elapsed time: " << difftime(end, start) << " seconds" << endl;
	log.close();

	return 0;
}

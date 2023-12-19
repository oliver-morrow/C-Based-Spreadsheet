// Name: Oliver Morrow
// Date: 12/06/2023
// Project: Excel Spreadsheet, ELEC 278 F23
// Queen's University, Smith Engineering ECE

#include "model.h"
#include "interface.h"
#include <stddef.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <ctype.h>
#include <stdbool.h>

// Defines a struct called Node, which represents a node in a linked list.
// Each node can have a type of either REFERENCE or CONSTANT.
// If the type is REFERENCE, the content of the node is a cell reference, consisting of a row and column index.
// If the type is CONSTANT, the content of the node is a constant value.
// The struct also contains a pointer to the next node in the list.
typedef struct Node {
    enum { REFERENCE, CONSTANT } type;
    union {
        struct {
            int row;
            int col;
        } reference; // Cell reference
        double constant; // Constant value
    } content;
    struct Node *next; // Pointer to the next node in the list
} Node;

// This code defines a struct called Cell, which represents a cell in a spreadsheet.
// Each cell can have a type of TEXT, NUMBER, FORMULA, or BLANK.
// If the type is TEXT, the content of the cell is a string.
// If the type is NUMBER, the content of the cell is a numeric value.
// If the type is FORMULA, the content of the cell is a linked list of nodes representing a parsed formula.
// If the type is BLANK, the cell is empty.
// The struct also contains additional fields such as the original formula string, an array of pointers to cells that depend on this cell, and the number of dependents.
typedef struct Cell {
    enum { TEXT, NUMBER, FORMULA, BLANK } type;
    union {
        char* text;      // For text and original formula string
        double number;   // For numeric values
        Node* formula;   // For parsed formula
    } content;
    char* original_formula; // Additional field to store the original formula string
    struct Cell **dependents; // Array of pointers to cells that depend on this cell
    int num_dependents;
} Cell;

// Initialize a spreadsheet using the structs above
Cell spreadsheet[NUM_ROWS][NUM_COLS];

// Convert a column letter to a column index (used for formula parsing)
int col_letter_to_index(char col_letter) {
    return toupper(col_letter) - 'A'; 
}

// Parse a formula string and return a linked list of nodes
// Each node is a structure with a type and a union
// The type can be CELL_REF or CONSTANT
// The union contains the content of the node, which is either a cell reference or a constant
Node *parse_formula(const char *formula_string, ROW current_row, COL current_col) {
    Node *head = NULL; // Initialize a pointer to the head of the linked list
    Node **current = &head; // Initialize a pointer to the current node

    const char *ptr = formula_string; // Initialize a pointer to the formula string

    // Loop through each character in the formula string
    while (*ptr) {
        if (*ptr == '=') { 
            ptr++; 
            // Using isalpha to check for alphabetic characters
        } else if (isalpha(*ptr)) {
            // Convert the letter to a column index
            int col = col_letter_to_index(*ptr);
            ptr++; // Move to the next character

            int row = atoi(ptr) - 1; // Convert the number after the letter to a row index using atoi
            while (isdigit(*ptr)) ptr++; // Move to the next character until a non-digit character is encountered

            *current = malloc(sizeof(Node)); // Allocate memory for a new node
            // Checking if memory allocation failed
            if (*current == NULL) {
                fprintf(stderr, "Error: Failed to allocate memory for node\n"); 
                return NULL; 
            }

            (*current)->type = REFERENCE; // Set the node type to REFERENCE
            (*current)->content.reference.row = row; // Set the row index of the reference
            (*current)->content.reference.col = col; // Set the column index of the reference
            (*current)->next = NULL; // Set the next pointer of the node to NULL
            current = &((*current)->next); // Move the current pointer to the next node

            Cell *referenced_cell = &spreadsheet[row][col]; // Get a pointer to the referenced cell in the spreadsheet
            referenced_cell->dependents = realloc(referenced_cell->dependents, (referenced_cell->num_dependents + 1) * sizeof(Cell *)); // Reallocate memory for the dependents array of the referenced cell
            // Check if memory reallocation failed
            if (referenced_cell->dependents == NULL) {
                fprintf(stderr, "Error: Failed to reallocate memory for dependents array\n"); 
                return NULL; 
            }

            referenced_cell->dependents[referenced_cell->num_dependents] = &spreadsheet[current_row][current_col]; // Add the current cell as a dependent of the referenced cell
            referenced_cell->num_dependents++; // Increment the number of dependents for the referenced cell
        } 
        // Check if the character is a digit or a decimal point
        else if (isdigit(*ptr) || *ptr == '.') {
            double constant = strtod(ptr, (char **)&ptr); // Convert the substring starting from the current character to a double constant
            *current = malloc(sizeof(Node)); // Allocate memory for a new node
            if (*current == NULL) { 
                fprintf(stderr, "Error: Failed to allocate memory for node\n"); 
                return NULL; 
            }

            (*current)->type = CONSTANT; // Set the node type to CONSTANT
            (*current)->content.constant = constant; // Set the constant value of the node
            (*current)->next = NULL; // Set the next pointer of the node to NULL
            current = &((*current)->next); // Move the current pointer to the next node
        } else {
            ptr++; 
        }
    }

    return head; // Return the head of the linked list. Yay!
}

// Checks if a given cell is involved in a circular dependency in a spreadsheet.
// If it finds that the cell in question is already in the recalculation array,
// it means that while recalculating the value of this cell, we have encountered this cell again, 
// indicating a circular dependency. In this case, it returns true.
// If it iterates over the entire recalculation array without finding the cell, 
// it means there is no circular dependency involving this cell, and it returns false.
// FUNCTION NOT USED IN THE FINAL VERSION OF THE PROGRAM since it does not work!
// I think 
bool is_circular_dependency(Cell *cell, Cell **recalculation_stack, int stack_size) {
    // Check if the current cell is already in the recalculation stack
    for (int i = 0; i < stack_size; i++) {
        if (recalculation_stack[i] == cell) {
            return true; // Circular dependency detected
        }
    }
    // If the cell is of type FORMULA, recursively check for circular dependency in its references
    // Recursive was good for this because it can be n amounts of references, technically
    if (cell->type == FORMULA && cell->content.formula != NULL) {
        // Creating a pointer to the formula with the formula linked list node
        Node *current = cell->content.formula;
        while (current != NULL) {
            if (current->type == REFERENCE) {
                Cell *referenced_cell = &spreadsheet[current->content.reference.row][current->content.reference.col];
                if (is_circular_dependency(referenced_cell, recalculation_stack, stack_size + 1)) {
                    return true;
                }
            }
            current = current->next;
        }
    }

    return false;
}


// Evaluates a formula represented by a linked list of nodes.
// The formula can contain cell references and constants.
// If a cell reference is encountered, it retrieves the value from the corresponding cell in the spreadsheet.
// If a constant is encountered, it adds the constant value to the result.
// If an invalid cell reference is encountered, it prints an error message and returns 0.0.
// If a non-numeric cell is encountered, it prints a warning message and treats the cell as 0 in the formula.
double evaluate_formula(Node *formula) {
    double result = 0.0; // Initialize the result to 0.0
    while (formula != NULL) { // Iterate through the formula linked list
        // If the node represents a cell reference
        if (formula->type == REFERENCE) {
            if (formula->content.reference.row < 0 || formula->content.reference.row >= NUM_ROWS || 
                formula->content.reference.col < 0 || formula->content.reference.col >= NUM_COLS) {
                // Check if the cell reference is valid
                fprintf(stderr, "Error: Invalid cell reference in formula\n"); // Print an error message
                return 0.0; // Return 0.0 (double) as the result
            }
            Cell cell = spreadsheet[formula->content.reference.row][formula->content.reference.col]; // Retrieve the cell from the spreadsheet
            if (cell.type == NUMBER) {
                result += cell.content.number;
            } else if (cell.type == FORMULA){
                // Check if the cell is involved in a circular dependency
                double formula_result = evaluate_formula(cell.content.formula); // Recursively evaluate the formula
                result += formula_result; // Add the result of the formula to the result
            } else {
                // If the cell is not a number or a formula, treat it as 0 and warn the user
                fprintf(stderr, "Warning: Non-numeric cell at [%d, %d] treated as 0 in formula\n", formula->content.reference.row, formula->content.reference.col); // Print a warning message
            }
        } else if (formula->type == CONSTANT) {
            // If the node represents a constant, add the constant value to the result
            result += formula->content.constant;
        }
        formula = formula->next; // Move to the next node in the formula linked list
    }
    return result; // Return the final result of the formula
}

// Update the dependents of a cell
// This function is called when a cell is updated
void update_dependents(ROW row, COL col, Cell **recalculation, int size) {
    // Check if the row and col are within the valid range
    if (row < 0 || row >= NUM_ROWS || col < 0 || col >= NUM_COLS) {
        update_cell_display(row, col, "INVALID CELL");
        return;
    }
    Cell *cell = &spreadsheet[row][col];
    // Check for circular dependency
    for (int i = 0; i < size; i++) {
        if (recalculation[i] == cell) {
            // Supposed to update the cell display to show circular error message
            // Could not get this function to work properly
            is_circular_dependency(cell, recalculation, size);
            update_cell_display(row, col, "CIRCULAR ERROR");
            return;
        }
    }
    // Check if the cell is valid
    if (cell == NULL) {
        update_cell_display(row, col, "INVALID CELL"); // Update the cell display to show the error message
        return;
    }

    recalculation[size] = cell;

    // Iterate through the dependents of the cell
    for (int i = 0; i < cell->num_dependents; i++) {
        // Get a pointer to the dependent cell with type Cell struct 
        Cell *dependent = cell->dependents[i];
        int dependent_row, dependent_col;

        // Find the row and column of the dependent cell in the spreadsheet
        for (dependent_row = 0; dependent_row < NUM_ROWS; dependent_row++) {
            for (dependent_col = 0; dependent_col < NUM_COLS; dependent_col++) {
                if (&spreadsheet[dependent_row][dependent_col] == dependent) {
                    break;
                }
            }
            if (dependent_col < NUM_COLS) {
                break;
            }
        }

        // Recalculate the value of the dependent cell if it contains a formula
        if (dependent->type == FORMULA) {
            double formula_result = evaluate_formula(dependent->content.formula);
            dependent->content.number = formula_result; // Update the number, keep the type as FORMULA

            // Update the display of the dependent cell with the new value
            char result_str[64];
            // Format the result as a string with one decimal place (VERY IMPORTANT!)
            snprintf(result_str, sizeof(result_str), "%.1f", formula_result);
            update_cell_display(dependent_row, dependent_col, result_str);
        }

        // Recursively update the dependents of the dependent cell
        if (size < 70) {
            // If the array size is less than 70, call the function recursively
            update_dependents(dependent_row, dependent_col, recalculation, (size + 1));
        } else {
            // If the array size is too large, print an error message and return
            fprintf(stderr, "Error: Size too large\n");
            return;
        }
    }
}

// Initialize the spreadsheet
void model_init() {
    for (int i = 0; i < NUM_ROWS; i++) {
        for (int j = 0; j < NUM_COLS; j++) {
            spreadsheet[i][j].type = BLANK; // Set cell type to BLANK
            spreadsheet[i][j].content.text = NULL; // Set content text to NULL
            spreadsheet[i][j].original_formula = NULL; // Set original formula to NULL
            spreadsheet[i][j].dependents = NULL; // Set dependents to NULL
            spreadsheet[i][j].num_dependents = 0; // Set number of dependents to 0
        }
    }
}

// Helper function to free memory for a formula linked list
// It takes in a pointer to the head of the linked list as a parameter
void free_formula(Node *formula) {
    while (formula != NULL) {
        Node *temp = formula;
        formula = formula->next;
        free(temp);
    }
}

// Function to update the value of a cell
void set_cell_value(ROW row, COL col, char *text) {
    // Handle the NULL case for text
    if (text == NULL) {
        clear_cell(row, col);
        return;
    }

    // Check if the text is a formula 
    if (text[0] == '=') {
        // Free existing memory if there is already a formula in the cell
        if (spreadsheet[row][col].type == FORMULA) {
            free_formula(spreadsheet[row][col].content.formula); // Free memory for the formula linked list
            free(spreadsheet[row][col].original_formula); // Free memory for the original formula string
        }
        // Store the original formula string
        spreadsheet[row][col].original_formula = strdup(text);

        // Parse, evaluate and update display for the formula
        Node *formula = parse_formula(text, row, col);
        spreadsheet[row][col].type = FORMULA; // Set the cell type to FORMULA
        spreadsheet[row][col].content.formula = formula; // Store the parsed formula

        // Evaluate the formula and update the display
        double formula_result = evaluate_formula(formula);
        char result_str[64];
        // Format the result as a string with one decimal place (VERY IMPORTANT for testing!)
        snprintf(result_str, sizeof(result_str), "%.1f", formula_result);
        update_cell_display(row, col, result_str);
    } else {
        char *endptr;
        // strtod converts a string to a double
        double number = strtod(text, &endptr);

        // Check if the entire string was a valid number
        if (*endptr == '\0') {
            // It's a number
            spreadsheet[row][col].type = NUMBER;
            spreadsheet[row][col].content.number = number;
            char number_str[64];
            snprintf(number_str, sizeof(number_str), "%.1f", number);
            update_cell_display(row, col, number_str);
        } else {
            // It's text
            // Free existing memory if there is already text in the cell
            if (spreadsheet[row][col].type == TEXT && spreadsheet[row][col].content.text != NULL) {
                free(spreadsheet[row][col].content.text);
            }

            // Allocate memory for new text and copy it
            spreadsheet[row][col].content.text = malloc(strlen(text) + 1);
            strcpy(spreadsheet[row][col].content.text, text);
            if (spreadsheet[row][col].content.text == NULL) {
                fprintf(stderr, "Memory allocation failed for cell text\n");
                exit(1);
            }
            // Set the cell type to TEXT
            spreadsheet[row][col].type = TEXT;
            update_cell_display(row, col, text);
        }
    }
    // Update the dependents of the cell
    Cell *recalculation[70]; // Maximum number of cells in the spreadsheet will always be 70 since it is fixed
    update_dependents(row, col, recalculation, 0); // We start with an empty array and update it through recursion
}

// Free memory for the cell and reset it to type BLANK and text NULL
void clear_cell(ROW row, COL col) {
    // Determine if the cell is valid
    if (row < 0 || row >= NUM_ROWS || col < 0 || col >= NUM_COLS) {
        return;
    }
    
    Cell *cell = &spreadsheet[row][col];

    // Free memory based on the type of the cell and reset it
    if (cell->type == TEXT && cell->content.text != NULL) {
        free(cell->content.text);
    } else if (cell->type == FORMULA && cell->content.formula != NULL) {
        free_formula(cell->content.formula);
        free(cell->original_formula);
    }

    // Reset the cell
    cell->type = BLANK;
    cell->content.text = NULL; // Applicable to both TEXT and FORMULA
    cell->num_dependents = 0; // Reset the number of dependents
    cell->original_formula = NULL; // Reset the original formula
}

// Function to retrieve the textual value of a cell
// Returns a string representing the value of the cell at the given coordinates
// It takes in the row and column of the cell as parameters.
// It returns a string representing the value of the cell at the given coordinates.
// If the cell coordinates are invalid, an error message is returned.
char *get_textual_value(ROW row, COL col) {
    // Check for invalid cell coordinates
    if (row < 0 || row >= NUM_ROWS || col < 0 || col >= NUM_COLS) {
        // Allocate memory for the error message
        char *errorMessage = malloc(100 * sizeof(char));
        if (errorMessage == NULL) {
            return strdup("Error: Memory allocation failed"); // Fallback error message if memory allocation fails
        }
        // Format the error message with the invalid coordinates
        snprintf(errorMessage, 100, "Error: Invalid cell coordinates [%d, %d]", row, col);
        return errorMessage;
    }
    // Access the cell from the spreadsheet using the provided row and column indices
    Cell *cell = &(spreadsheet[row][col]);
    char *result;
    
    // Switch case to handle different types of cell content
    switch (cell->type) {
        case TEXT:
            // Check if the text cell is empty
            if (cell->content.text != NULL) {
                result = strdup(cell->content.text); // Duplicate the text value
            } else {
                result = strdup("Error: Empty text cell"); // Return an error message for an empty text cell
            }
            break;
        case NUMBER:
            // Allocate memory for the result string
            result = malloc(64 * sizeof(char));
            if (result != NULL) {
                // Format the number value as a string
                snprintf(result, 64, "%f", cell->content.number);
            } else {
                return strdup("Error: Memory allocation failed"); // Return an error message if memory allocation fails
            }
            break;
        case FORMULA:
            // Check if the formula cell has an original formula
            if (cell->original_formula != NULL) {
                result = strdup(cell->original_formula); // Duplicate the original formula string
            } else {
                result = strdup("Error: Empty formula cell"); // Return an error message for an empty formula cell
            }
            break;
        case BLANK:
            result = strdup(""); // Return an empty string for a blank cell
            break;
        default:
            result = strdup("Error: Unknown cell type"); // Return an error message for an unknown cell type
            break;
    }

    if (result == NULL) {
        result = strdup("Error: Unknown cell content"); // Return an error message for undefined cell content
    }
    return result;
}


// Have a good holiday break!
// I think the method I picked was too complicated and when the formulas came in i should have switched
// Long functions became a bit hard to understand and I think I should have used more helper functions
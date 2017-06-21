# -*- coding: cp1252 -*-
from xlutils.copy import copy
from xlwt import easyxf
from xlrd import open_workbook

class CNA():
        def __init__(self, txt_name, xls_name, xls_final_name, num_rows, num_cols):
                self.txt_name = txt_name
                self.xls_name = xls_name
                self.xls_final_name = xls_final_name
                self.num_rows = num_rows
                self.num_cols = num_cols
                
        def mat1(self):
                """Reads initial matrix from a text file
                """
                f = open(self.txt_name, 'r')
                mat = []
                for line in f:
                        mat.append(line.strip().split())
                for i in range(len(mat)):
                        for j in range(len(mat[i])):
                                mat[i][j] = int(mat[i][j])
                return mat

        def get_matrix_info(self):
                """Gets the number of machines and products in the given matrix
                """
                rb = open_workbook(self.xls_name)
                sheet = rb.sheet_by_index(0)
                products = []
                machines = []
                for i in range(self.num_rows):
                        machines.append(sheet.cell_value(i + 2, 0))
                for j in range(self.num_cols):
                        products.append(sheet.cell_value(0, j + 2))
                return [machines, products]
                
        def mat1_xls(self):
                """Reads initial matrix from a xls worksheet
                """
                rb = open_workbook(self.xls_name)
                sheet = rb.sheet_by_index(0)
                mat = []
                for i in range(self.num_rows):
                    mat.append([0]*self.num_cols)
                for i in range(self.num_rows):
                    for j in range(self.num_cols):
                                mat[i][j] = sheet.cell_value(i+2, j+2)
                return mat

        def mat2_xls(self, mat2):
                """Writes the resulting matrix in a xls file
                """
                rb = open_workbook(self.xls_final_name, encoding_override="cp1252", formatting_info=True)
                wb = copy(rb)
                sheet = wb.get_sheet(1)
                style = easyxf('font: name Arial, color black; align: wrap on, vert centre, horiz center; borders: left thin, right thin, top thin, bottom thin;')
                machines, products = self.get_matrix_info()
                sheet.write(1, 1, "Indice", style)
                for i in range(len(machines)): #Writing the machine's names on the workbook
                        sheet.write(i + 2, 0, machines[i])
                for i in range(len(machines)): #Writing the machine's names on the workbook
                        sheet.write(i + 2, 1, i + 1)                
                for j in range(len(machines)): #Writing the machine's names on the workbook
                        sheet.write(0, j + 2, machines[j])
                for j in range(len(machines)): #Writing the machine's names on the workbook
                        sheet.write(1, j + 2, j + 1)                
                for i in range(len(mat2)):
                        for j in range(len(mat2[i])):
                                sheet.write(i + 2, j + 2, mat2[i][j], style)
                wb.save(self.xls_final_name)

        def mat3_xls(self, mat3, ordemlinhas):
                """Writes the resulting matrix in a xls file
                """
                rb = open_workbook(self.xls_final_name, encoding_override="cp1252", formatting_info=True)
                wb = copy(rb)
                sheet = wb.get_sheet(2)
                style = easyxf('font: name Arial, color black; align: wrap on, vert centre, horiz centre; borders: left thin, right thin, top thin, bottom thin;')
                machines, products = self.get_matrix_info()
                sheet.write(1, 1, "Indice", style)
                for j in range(self.num_cols): #Writing the product's names on the workbook
                        sheet.write(0, j + 2, products[j])
                for j in range(self.num_cols): #Writing the product's names on the workbook
                        sheet.write(1, j + 2, j + 1)
                for i in range(len(ordemlinhas)): #Writing the machine's names on the workbook
                        sheet.write(i + 2, 0, machines[ordemlinhas[i]])
                for i in range(len(ordemlinhas)): #Writing the machine's names on the workbook
                        sheet.write(i + 2, 1, ordemlinhas[i] + 1)
                for i in range(len(mat3)):
                        for j in range(len(mat3[i])):
                                sheet.write(i + 2, j + 2, mat3[i][j], style)
                wb.save(self.xls_final_name)

        def mat4_xls(self, mat4, index_mach, index_prod):
                """Writes the resulting matrix in a xls file
                """
                rb = open_workbook(self.xls_final_name, encoding_override="cp1252", formatting_info=True)
                wb = copy(rb)
                sheet = wb.get_sheet(3)
                style = easyxf('font: name Arial, color black; align: wrap on, vert centre, horiz center; borders: left thin, right thin, top thin, bottom thin;')
                machines, products = self.get_matrix_info()
                for j in range(len(mat4[0])): #Writing the product's names on the workbook
                        sheet.write(0, j + 2, products[index_prod[j]])
                for i in range(len(mat4)): #Writing the machine's names on the workbook
                        sheet.write(i + 2, 0, machines[index_mach[i]])
                for i in range(len(mat4)):
                        for j in range(len(mat4[i])):
                                sheet.write(i + 2, j + 2, mat4[i][j], style)
                wb.save(self.xls_final_name)

        def mat2(self):
                """Calculates the matrix 2, based on the initial matrix
                """
                mat = self.mat1_xls()
                tasks_per_machine = [0]*len(mat)
                mat2 = []
                for _ in range(len(mat)):
                        mat2.append([0]*len(mat))
                for i in range((len(mat) - 1)):
                        for m in range(i + 1, len(mat)):
                                for j in range(len(mat[i])):
                                                if mat[i][j] == 1 and mat[m][j] == 1:
                                                        tasks_per_machine[i] += 1
                                                        tasks_per_machine[m] += 1
                                                        mat2[i][m] += 1
                                                        mat2[m][i] += 1
 #               self.mat2_xls(mat2)
                return [mat2, tasks_per_machine]

        def mat3(self, matriz1, matriz2, tasks_per_machine):
                """Calculates the matrix 3, based on the matrix 2.
                """
                ordemlinhas = []
                #Primeira Linha
                maior = 0
                for i in xrange(1,len(tasks_per_machine)):
                        if tasks_per_machine[i] > tasks_per_machine[maior]:
                                maior = i
                ordemlinhas.append(maior)
                #Restante das Linhas
                while(len(ordemlinhas)<len(tasks_per_machine)):
                        i = ordemlinhas[-1]
                        for j in xrange(len(tasks_per_machine)):
                                if j not in ordemlinhas:
                                        maior = j
                                        break
                        for j in xrange(len(tasks_per_machine)):
                                if j not in ordemlinhas:
                                        if matriz2[i][j] > matriz2[i][maior]:
                                                maior = j
                                        elif matriz2[i][j] == matriz2[i][maior]:
                                                if tasks_per_machine[j] > tasks_per_machine[maior]:
                                                        maior = j
                        ordemlinhas.append(maior)
                #Construcao da nova matriz com as linhas reordenadas2
                novaMat = []
                for i in ordemlinhas:
                        novaMat.append(matriz1[i])
                return [novaMat, ordemlinhas]

        def contagem(self, seq, coluna):
                """Returns the quantity of numbers '1' in a given col
                """
                cont = 0
                for linha in seq:
                        if linha[coluna] == 1:
                                cont += 1
                return cont

        def mat4(self, matriz3):
                """Calculates the matrix 4, based on the matrix 3
                """
                mat = matriz3[:]
                ordemColunas = []
                colunas = len(mat[0])
                while True:
                        #Se a matriz ainda esta divisivel
                        if len(mat) > 1:
                                div = len(mat)/2
                                seq1 = mat[:div]
                                seq2 = mat [div:]
                                #Adiciona a coluna na nova ordem se o numero de 1s
                                #na coluna for maior em SEQ1 que em SEQ2
                                for j in xrange(colunas):
                                        if j not in ordemColunas:
                                                n1 = self.contagem(seq1,j)
                                                n2 = self.contagem(seq2,j)
                                                if n1 > n2:
                                                        ordemColunas.append(j)
                                #Se todas as linhas ja foram marcadas o processo para,
                                #caso contrario a nova matriz sera SEQ2
                                if len(ordemColunas) == colunas:
                                        break
                                else:
                                        mat = seq2
                        #Caso nao seja mais possivel dividir a matriz
                        #adicionam-se as linhas restantes diretamente
                        else:
                                for j in xrange(colunas):
                                        if j not in ordemColunas:
                                                ordemColunas.append(j)
                                break
                #Construcao da nova matriz, com linhas e colunas rearranjadas
                novaMat = []
                for i in matriz3:
                        linha = []
                        for j in ordemColunas:
                                linha.append(i[j])
                        novaMat.append(linha)
                return [novaMat, ordemColunas]

if __name__ == "__main__":
        filename = "matriz1.txt"
        filename_xls = "Analise de Fluxo Produtos e Máquinas Final.xls"
        final_xls = "Matriz.xls"
        cna = CNA(filename, filename_xls, final_xls, 27, 30)
        matriz1 = cna.mat1_xls()
        print 'Matriz A'
        for i in matriz1:
        	print i
        matriz2, tasks_per_machine = cna.mat2()
        print
        for i in xrange(len(tasks_per_machine)):
        	matriz2[i].append('|')
        	matriz2[i].append(tasks_per_machine[i])
        print "Matriz B"
        for i in matriz2:
        	print i
        print
        nova, rows = cna.mat3(matriz1, matriz2, tasks_per_machine)
        print "Ordem das linhas\n"
        print rows
        print
        print 'Matriz A com as linhas atualizadas'
        for i in nova:
        	print i
        print
        mat4, cols = cna.mat4(nova)
        print "Ordem das colunas\n"
        print cols
        print
        for i in mat4:
                print i
        cna.mat4_xls(mat4, rows, cols)

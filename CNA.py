from xlutils.copy import copy
from xlwt import easyxf
from xlrd import open_workbook

def mat1(filename):
        f = open(filename, 'r')
        mat = []
        for line in f:
                mat.append(line.strip().split())
        for i in range(len(mat)):
                for j in range(len(mat[i])):
                        mat[i][j] = int(mat[i][j])
        return mat

def get_matrix_info(filename):
        rb = open_workbook(filename)
        sheet = rb.sheet_by_index(0)
        products = []
        machines = []
        for i in range(27):
                machines.append(sheet.cell_value(i + 2, 0))
        for j in range(28):
                products.append(sheet.cell_value(0, j + 2))
        return [machines, products]
        
def mat1_xls(filename):
        rb = open_workbook(filename)
        sheet = rb.sheet_by_index(0)
        mat = []
        for i in range(27):
            mat.append([0]*28)
        for i in range(27):
            for j in range(28):
                        mat[i][j] = sheet.cell_value(i+2, j+2)
        return mat

def mat4_xls(filename, mat4, index_mach, index_prod):
        rb = open_workbook(filename, encoding_override="cp1252", formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        style = easyxf('font: name calibri, color black; align: wrap on, vert centre, horiz center; borders: left thin, right thin, top thin, bottom thin;')
        machines, products = get_matrix_info(filename)
        for j in range(len(mat4[0])): #Writing the product's names on the workbook
                sheet.write(0, j + 2, products[index_prod[j]])
        for i in range(len(mat4)): #Writing the machine's names on the workbook
                sheet.write(i + 2, 0, machines[index_mach[i]])
        for i in range(len(mat4)):
                for j in range(len(mat4[i])):
                        sheet.write(i + 2, j + 2, mat4[i][j], style)
        wb.save(filename)

def mat2(filename):
        mat = mat1_xls(filename)
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
        return [mat2, tasks_per_machine]

def mat3(matriz1,matriz2,tasks_per_machine):
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

def contagem(seq, coluna):
        cont = 0
        for linha in seq:
                if linha[coluna] == 1:
                        cont += 1
        return cont

def mat4(matriz3):
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
                                        n1 = contagem(seq1,j)
                                        n2 = contagem(seq2,j)
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
        filename_xls = "Matriz.xls"
        matriz1 = mat1_xls(filename_xls)
        print 'Matriz A'
        for i in matriz1:
        	print i
        matriz2, tasks_per_machine = mat2(filename_xls)
        print
        for i in xrange(len(tasks_per_machine)):
        	matriz2[i].append('|')
        	matriz2[i].append(tasks_per_machine[i])
        print "Matriz B"
        for i in matriz2:
        	print i
        print
        nova, rows = mat3(matriz1,matriz2,tasks_per_machine)
        print "Ordem das linhas\n"
        print rows
        print
        print 'Matriz A com as linhas atualizadas'
        for i in nova:
        	print i
        print
        mat4, cols = mat4(nova)
        print "Ordem das colunas\n"
        print cols
        print
        for i in mat4:
                print i
        mat4_xls("Matriz Final.xls", mat4, rows, cols)

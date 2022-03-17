# Extrator de dados de PDF

---

Utilizando a library *pumupdf* esse script extrai dados do pdf desejado, e com a library *openpyxl* insere esses mesmos dados em uma planilha.


Foi utilizado a library OS para poder percorrer uma pasta (*file_path*), buscando os arquivos PDF. Cada arquivo encontrado, é enviado para a função *extract_data*  


Com a função abaixo, foi extraido todo o texto do PDF, fiz isso para ver o indice dos dados desejados e em seguida, colocar esses dados em uma lista e passar para a proxima função 

![codesnippet](https://user-images.githubusercontent.com/68126144/158893394-8e9a892e-7f87-4d1c-82a3-0ab28326ee70.png)


A função *store_data* recebe como parametro os daddos extraidos e os insere em uma nova linha da planilha (que já deve estar criada e passada o caminho dentro da variavel *excel_path*)

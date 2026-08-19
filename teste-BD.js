const XLSX = require("xlsx");
const arquivo = XLSX.readFile("dados-exemplo.xlsx");
const planilha = arquivo.Sheets["dados"];
const dados = XLSX.utils.sheet_to_json(planilha);
const readline = require("readline");
const entrada = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

console.log("mande 'chamar' para chamar pelo ID, 'add' para adicionar dados extras e delete para apagar.")

entrada.on("line", function(texto){

    if (texto === "chamar") {
        entrada.question("digite o ID: ", function(id) {
            chamarComID(Number(id));
        });
    };

    if (texto === "add") {
        entrada.question("digite o nome: ", function (nome) {
            entrada.question("digite o idade: ", function (idade) {
                entrada.question("digite o curso: ", function (curso) {
                        
                   adicionarDados(
                        nome,
                        Number(idade),
                        curso
                    );
                });
            });
        });
    };
    if (texto === "delete") {
        entrada.question("Digite o ID: ", function(id) {
            apagarComID(Number(id));
        });
    }
}); 

function chamarComID(id) {
    const pessoa = dados.find(pessoa => pessoa.ID === id);

    console.log(pessoa);
};

function adicionarDados(nome, idade, curso) {

    let novoID;

    if (dados.length === 0) {
        novoID = 1;
    } else {
        const maiorID = Math.max(...dados.map(pessoa => pessoa.ID));
        novoID = maiorID + 1;
    }

    const pessoa = {
        ID: novoID,
        Nome: nome,
        Idade: idade,
        Curso: curso
    };

    dados.push(pessoa);

    dados.sort((a, b) => a.ID - b.ID);

    const novaplanilha = XLSX.utils.json_to_sheet(dados);
    arquivo.Sheets["dados"] = novaplanilha;

    XLSX.writeFile(arquivo, "dados-exemplo.xlsx");

    console.log("Pessoa adicionada!");
    console.log(pessoa);
};

function apagarComID(id) {

    const indice = dados.findIndex(pessoa => pessoa.ID === id);

    if (indice === -1) {
        console.log("ID não encontrado!");
        return;
    }

    dados.splice(indice, 1);

    const novaplanilha = XLSX.utils.json_to_sheet(dados);
    arquivo.Sheets["dados"] = novaplanilha;

    XLSX.writeFile(arquivo, "dados-exemplo.xlsx");

    console.log("Pessoa apagada com sucesso!");
};
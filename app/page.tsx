'use client';

import { useRef, useState, type ChangeEvent } from 'react';
import * as XLSX from 'xlsx';

type ValorCelula = string | number | boolean | Date | null | undefined;
type LinhaPlanilha = Record<string, ValorCelula>;

interface Infracao {
  veiculo: string;
  motorista: string;
  endereco: string;
  inicioData: Date;
  inicioString: string;
  fimData: Date;
  fimString: string;
  duracaoMin: string;
  maxVel: number;
}

interface RegistoRelatorio {
  dataObj: Date;
  dataString: string;
  veiculo: string;
  velocidade: number;
  motorista: string;
  endereco: string;
}

interface BlocoVelocidade {
  veiculo: string;
  motorista: string;
  endereco: string;
  inicioObj: Date;
  inicioStr: string;
  fimObj: Date;
  fimStr: string;
  maxVel: number;
}

interface MapeamentoColunas {
  data: string;
  veiculo: string;
  velocidade: string;
  motorista?: string;
  endereco?: string;
}

const LIMITE_VELOCIDADE = 130;
const DURACAO_MINIMA_MINUTOS = 1;

const normalizarCabecalho = (valor: string): string =>
  valor
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

const limparTexto = (valor: ValorCelula): string => {
  if (valor === null || valor === undefined) {
    return '';
  }

  return String(valor).trim();
};

const converterNumero = (valor: ValorCelula): number | null => {
  if (typeof valor === 'number' && Number.isFinite(valor)) {
    return valor;
  }

  const texto = limparTexto(valor);

  if (!texto) {
    return null;
  }

  let normalizado = texto.replace(/\s/g, '');

  if (normalizado.includes(',') && normalizado.includes('.')) {
    normalizado = normalizado.replace(/\./g, '').replace(',', '.');
  } else if (normalizado.includes(',')) {
    normalizado = normalizado.replace(',', '.');
  }

  const numero = Number(normalizado);
  return Number.isFinite(numero) ? numero : null;
};

const converterDataExcel = (serial: number): Date => {
  const diasInteiros = Math.floor(serial);
  const fracaoDia = serial - diasInteiros;
  const data = new Date(1899, 11, 30);

  data.setDate(data.getDate() + diasInteiros);

  const totalSegundos = Math.round(fracaoDia * 24 * 60 * 60);
  const horas = Math.floor(totalSegundos / 3600);
  const minutos = Math.floor((totalSegundos % 3600) / 60);
  const segundos = totalSegundos % 60;

  data.setHours(horas, minutos, segundos, 0);
  return data;
};

const converterData = (valor: ValorCelula): Date | null => {
  if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
    return valor;
  }

  if (typeof valor === 'number' && Number.isFinite(valor)) {
    return converterDataExcel(valor);
  }

  const texto = limparTexto(valor);

  if (!texto) {
    return null;
  }

  const correspondenciaBrasileira = texto.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/,
  );

  if (correspondenciaBrasileira) {
    const [, diaTexto, mesTexto, anoTexto, horaTexto, minutoTexto, segundoTexto] =
      correspondenciaBrasileira;
    const ano = anoTexto.length === 2 ? Number(`20${anoTexto}`) : Number(anoTexto);
    const data = new Date(
      ano,
      Number(mesTexto) - 1,
      Number(diaTexto),
      Number(horaTexto ?? '0'),
      Number(minutoTexto ?? '0'),
      Number(segundoTexto ?? '0'),
    );

    return Number.isNaN(data.getTime()) ? null : data;
  }

  const dataPadrao = new Date(texto);
  return Number.isNaN(dataPadrao.getTime()) ? null : dataPadrao;
};

const preencher2 = (valor: number): string => String(valor).padStart(2, '0');

const formatarData = (data: Date): string =>
  `${preencher2(data.getDate())}/${preencher2(data.getMonth() + 1)}/${data.getFullYear()} ${preencher2(data.getHours())}:${preencher2(data.getMinutes())}:${preencher2(data.getSeconds())}`;

const formatarDataOriginal = (valor: ValorCelula, dataFallback: Date): string => {
  const texto = limparTexto(valor);
  return texto || formatarData(dataFallback);
};

const encontrarColuna = (cabecalhos: string[], candidatos: string[]): string | undefined => {
  const candidatosNormalizados = candidatos.map(normalizarCabecalho);

  return cabecalhos.find((cabecalho) =>
    candidatosNormalizados.includes(normalizarCabecalho(cabecalho)),
  );
};

const mapearColunas = (cabecalhos: string[]): MapeamentoColunas | null => {
  const data = encontrarColuna(cabecalhos, ['Data']);
  const veiculo = encontrarColuna(cabecalhos, ['Veículo', 'Veiculo']);
  const velocidade = encontrarColuna(cabecalhos, ['Velocidade']);

  if (!data || !veiculo || !velocidade) {
    return null;
  }

  return {
    data,
    veiculo,
    velocidade,
    motorista: encontrarColuna(cabecalhos, ['Motorista']),
    endereco: encontrarColuna(cabecalhos, ['Endereço', 'Endereco']),
  };
};

const finalizarBloco = (
  blocoAtual: BlocoVelocidade | null,
  listaInfracoes: Infracao[],
): void => {
  if (!blocoAtual) {
    return;
  }

  const duracaoMs = blocoAtual.fimObj.getTime() - blocoAtual.inicioObj.getTime();
  const duracaoMin = duracaoMs / (1000 * 60);

  if (duracaoMin >= DURACAO_MINIMA_MINUTOS) {
    listaInfracoes.push({
      veiculo: blocoAtual.veiculo,
      motorista: blocoAtual.motorista,
      endereco: blocoAtual.endereco,
      inicioData: blocoAtual.inicioObj,
      inicioString: blocoAtual.inicioStr,
      fimData: blocoAtual.fimObj,
      fimString: blocoAtual.fimStr,
      duracaoMin: duracaoMin.toFixed(1),
      maxVel: blocoAtual.maxVel,
    });
  }
};

export default function DetectorVelocidade() {
  const [infracoes, setInfracoes] = useState<Infracao[]>([]);
  const [ficheiroAtivo, setFicheiroAtivo] = useState<string | null>(null);
  const [abaAnalisada, setAbaAnalisada] = useState<string | null>(null);
  const [erro, setErro] = useState<string | null>(null);
  const [aProcessar, setAProcessar] = useState(false);
  const inputFicheiroRef = useRef<HTMLInputElement | null>(null);

  const limparConsulta = (): void => {
    setInfracoes([]);
    setFicheiroAtivo(null);
    setAbaAnalisada(null);
    setErro(null);
    setAProcessar(false);

    if (inputFicheiroRef.current) {
      inputFicheiroRef.current.value = '';
    }
  };

  const gerarRelatorioOcorrencias = (): void => {
    if (infracoes.length === 0) {
      return;
    }

    const dadosExportacao = infracoes.map((inf) => ({
      Veiculo: inf.veiculo,
      Motorista: inf.motorista,
      Inicio: inf.inicioString,
      Fim: inf.fimString,
      'Duracao (Minutos)': inf.duracaoMin,
      'Velocidade Maxima (km/h)': inf.maxVel,
      'Ultimo Endereco': inf.endereco,
    }));

    const worksheet = XLSX.utils.json_to_sheet(dadosExportacao);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Ocorrencias');

    const baseNome = (ficheiroAtivo ?? 'relatorio')
      .replace(/\.(xlsx|xls)$/i, '')
      .replace(/[\\/:*?"<>|]+/g, '-');

    XLSX.writeFile(workbook, `${baseNome}-ocorrencias.xlsx`);
  };

  const processarRelatorio = (dadosBinarios: ArrayBuffer): void => {
    try {
      const workbook = XLSX.read(dadosBinarios, {
        type: 'array',
        cellDates: true,
      });

      const nomeAba = workbook.SheetNames[0];

      if (!nomeAba) {
        setInfracoes([]);
        setErro('O ficheiro XLSX não contém nenhuma aba legível.');
        setAProcessar(false);
        return;
      }

      const worksheet = workbook.Sheets[nomeAba];
      const linhasBrutas = XLSX.utils.sheet_to_json<LinhaPlanilha>(worksheet, {
        defval: null,
        raw: true,
      });

      if (linhasBrutas.length === 0) {
        setInfracoes([]);
        setAbaAnalisada(nomeAba);
        setErro('A aba selecionada está vazia.');
        setAProcessar(false);
        return;
      }

      const cabecalhos = Object.keys(linhasBrutas[0]);
      const colunas = mapearColunas(cabecalhos);

      if (!colunas) {
        setInfracoes([]);
        setAbaAnalisada(nomeAba);
        setErro(
          'Não encontrei as colunas obrigatórias Data, Veículo e Velocidade na primeira aba.',
        );
        setAProcessar(false);
        return;
      }

      const dados: RegistoRelatorio[] = linhasBrutas
        .map((linha) => {
          const dataObj = converterData(linha[colunas.data]);
          const velocidade = converterNumero(linha[colunas.velocidade]);
          const veiculo = limparTexto(linha[colunas.veiculo]);

          if (!dataObj || velocidade === null || !veiculo) {
            return null;
          }

          return {
            dataObj,
            dataString: formatarDataOriginal(linha[colunas.data], dataObj),
            veiculo,
            velocidade,
            motorista: colunas.motorista
              ? limparTexto(linha[colunas.motorista]) || '---'
              : '---',
            endereco: colunas.endereco
              ? limparTexto(linha[colunas.endereco]) || '---'
              : '---',
          };
        })
        .filter((item): item is RegistoRelatorio => item !== null);

      if (dados.length === 0) {
        setInfracoes([]);
        setAbaAnalisada(nomeAba);
        setErro('Não encontrei linhas válidas para analisar neste relatório.');
        setAProcessar(false);
        return;
      }

      dados.sort((a, b) => {
        if (a.veiculo < b.veiculo) {
          return -1;
        }

        if (a.veiculo > b.veiculo) {
          return 1;
        }

        return a.dataObj.getTime() - b.dataObj.getTime();
      });

      const listaInfracoes: Infracao[] = [];
      let blocoAtual: BlocoVelocidade | null = null;

      for (let i = 0; i < dados.length; i += 1) {
        const atual = dados[i];

        if (atual.velocidade > LIMITE_VELOCIDADE) {
          if (!blocoAtual || blocoAtual.veiculo !== atual.veiculo) {
            finalizarBloco(blocoAtual, listaInfracoes);

            blocoAtual = {
              veiculo: atual.veiculo,
              motorista: atual.motorista,
              endereco: atual.endereco,
              inicioObj: atual.dataObj,
              inicioStr: atual.dataString,
              fimObj: atual.dataObj,
              fimStr: atual.dataString,
              maxVel: atual.velocidade,
            };
          } else {
            blocoAtual.fimObj = atual.dataObj;
            blocoAtual.fimStr = atual.dataString;
            blocoAtual.endereco = atual.endereco || blocoAtual.endereco;
            blocoAtual.motorista = atual.motorista || blocoAtual.motorista;

            if (atual.velocidade > blocoAtual.maxVel) {
              blocoAtual.maxVel = atual.velocidade;
            }
          }
        } else if (blocoAtual) {
          finalizarBloco(blocoAtual, listaInfracoes);
          blocoAtual = null;
        }
      }

      finalizarBloco(blocoAtual, listaInfracoes);

      setErro(null);
      setAbaAnalisada(nomeAba);
      setInfracoes(listaInfracoes);
      setAProcessar(false);
    } catch {
      setInfracoes([]);
      setErro('Não foi possível ler o ficheiro XLSX. Verifique se o relatório está válido.');
      setAProcessar(false);
    }
  };

  const gerirCarregamento = (e: ChangeEvent<HTMLInputElement>): void => {
    const ficheiro = e.target.files?.[0];

    if (!ficheiro) {
      return;
    }

    setFicheiroAtivo(ficheiro.name);
    setAbaAnalisada(null);
    setErro(null);
    setInfracoes([]);
    setAProcessar(true);

    const reader = new FileReader();

    reader.onload = (evento) => {
      const resultado = evento.target?.result;

      if (!(resultado instanceof ArrayBuffer)) {
        setErro('Falha ao carregar o ficheiro selecionado.');
        setAProcessar(false);
        return;
      }

      processarRelatorio(resultado);
    };

    reader.onerror = () => {
      setErro('Falha ao abrir o ficheiro selecionado.');
      setAProcessar(false);
    };

    reader.readAsArrayBuffer(ficheiro);
  };

  return (
    <div className="min-h-screen bg-gray-50 p-8">
      <div className="mx-auto max-w-5xl space-y-8">
        <div className="rounded-xl border border-gray-100 bg-white p-6 shadow-sm">
          <div className="space-y-3">
            <span className="inline-flex rounded-full bg-blue-100 px-3 py-1 text-xs font-semibold uppercase tracking-wide text-blue-700">
              SpeedMaster
            </span>
            <h1 className="text-2xl font-bold text-gray-800">
              Central de Análise de Velocidade
            </h1>
          </div>
          <p className="mt-2 text-gray-500">
            O SpeedMaster detecta ocorrências acima de 130 km/h em relatórios
            XLSX do Infleet com duração mínima de 1 minuto.
          </p>

          <div className="mt-6 flex w-full items-center justify-center">
            <label className="flex h-40 w-full cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed border-blue-300 bg-blue-50 transition-colors hover:bg-blue-100">
              <div className="flex flex-col items-center justify-center pt-5 pb-6">
                <svg
                  className="mb-3 h-10 w-10 text-blue-500"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                  xmlns="http://www.w3.org/2000/svg"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                  />
                </svg>
                <p className="mb-2 text-sm text-gray-700">
                  <span className="font-semibold">Clique para carregar</span> ou
                  arraste e solte o ficheiro XLSX
                </p>
                <p className="text-xs text-gray-500">
                  Apenas relatórios Infleet (.xlsx e .xls)
                </p>
              </div>
              <input
                ref={inputFicheiroRef}
                type="file"
                className="hidden"
                accept=".xlsx,.xls"
                onChange={gerirCarregamento}
              />
            </label>
          </div>

          {ficheiroAtivo && (
            <div className="mt-4 space-y-1 text-sm font-medium text-green-600">
              <p>Ficheiro carregado: {ficheiroAtivo}</p>
              {abaAnalisada && <p>Aba analisada: {abaAnalisada}</p>}
            </div>
          )}

          {(ficheiroAtivo || infracoes.length > 0 || erro) && (
            <div className="mt-6 flex flex-col gap-3 sm:flex-row">
              <button
                type="button"
                onClick={gerarRelatorioOcorrencias}
                disabled={infracoes.length === 0}
                className="inline-flex items-center justify-center rounded-lg bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition-colors hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-blue-300"
              >
                Gerar Relatório das Ocorrências
              </button>
              <button
                type="button"
                onClick={limparConsulta}
                className="inline-flex items-center justify-center rounded-lg border border-gray-200 bg-white px-4 py-2 text-sm font-semibold text-gray-700 transition-colors hover:bg-gray-50"
              >
                Nova Consulta
              </button>
            </div>
          )}
        </div>

        {aProcessar ? (
          <p className="text-center text-gray-500">A processar os dados...</p>
        ) : erro ? (
          <div className="rounded-xl border border-red-100 bg-white p-6 text-center shadow-sm">
            <p className="font-medium text-red-600">{erro}</p>
          </div>
        ) : infracoes.length > 0 ? (
          <div className="overflow-hidden rounded-xl border border-gray-100 bg-white shadow-sm">
            <div className="border-b border-gray-100 bg-gray-50/50 p-6">
              <h2 className="text-lg font-bold text-gray-800">
                Resultados Encontrados ({infracoes.length})
              </h2>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm text-gray-600">
                <thead className="border-b border-gray-100 bg-gray-50 text-xs uppercase text-gray-700">
                  <tr>
                    <th className="px-6 py-4">Veículo</th>
                    <th className="px-6 py-4">Motorista</th>
                    <th className="px-6 py-4">Início</th>
                    <th className="px-6 py-4">Fim</th>
                    <th className="px-6 py-4">Duração (Minutos)</th>
                    <th className="px-6 py-4 text-right">Velocidade Máxima</th>
                    <th className="px-6 py-4">Último Endereço</th>
                  </tr>
                </thead>
                <tbody>
                  {infracoes.map((inf) => (
                    <tr
                      key={`${inf.veiculo}-${inf.inicioString}-${inf.fimString}`}
                      className="border-b border-gray-50 bg-white transition-colors hover:bg-gray-50"
                    >
                      <td className="px-6 py-4 font-medium text-gray-900">
                        {inf.veiculo}
                      </td>
                      <td className="px-6 py-4">{inf.motorista}</td>
                      <td className="px-6 py-4">{inf.inicioString}</td>
                      <td className="px-6 py-4">{inf.fimString}</td>
                      <td className="px-6 py-4">
                        <span className="rounded bg-yellow-100 px-2.5 py-0.5 text-xs font-medium text-yellow-800">
                          {inf.duracaoMin} min
                        </span>
                      </td>
                      <td className="px-6 py-4 text-right font-bold text-red-600">
                        {inf.maxVel} km/h
                      </td>
                      <td className="px-6 py-4">{inf.endereco}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          ficheiroAtivo &&
          !aProcessar && (
            <div className="rounded-xl border border-gray-100 bg-white p-6 text-center shadow-sm">
              <p className="text-lg font-medium text-green-600">
                Nenhuma infração prolongada detetada neste relatório XLSX.
              </p>
            </div>
          )
        )}
      </div>
    </div>
  );
}

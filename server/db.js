import mongoose from 'mongoose';

const recordSchema = new mongoose.Schema({
  id: { type: String, required: true, unique: true },
  data: String,
  encarregado: String,
  os: String,
  tipoServico: String,
  regional: String,
  conformidade: String,
  descNaoConformidade: String,
  processo: String,
  nomeEletricistaLider: String,
  nomeEletricista: String,
  quemInspecionou: String,
  tipoInspecao: String,
  matriculaLider: String,
  matriculaEletricista: String,
  registroFoto: String,
  observacao: String,
  naoConformidade: String,
  regrasOuro: { type: Object, default: {} },
  regrasArquivos: { type: Object, default: {} }
}, { timestamps: true });

export const Record = mongoose.model('Record', recordSchema);

let connectPromise = null;

export async function connectDB() {
  if (mongoose.connection.readyState === 1) {
    return;
  }
  
  if (connectPromise) {
    return connectPromise;
  }
  
  connectPromise = (async () => {
    try {
      const mongoUri = process.env.MONGODB_URI || 'mongodb://localhost:27017/indica';
      await mongoose.connect(mongoUri, {
        connectTimeoutMS: 5000,
        socketTimeoutMS: 45000,
        serverSelectionTimeoutMS: 5000,
        retryWrites: true,
        w: 'majority'
      });
      console.log('MongoDB connected');
      connectPromise = null;
    } catch (err) {
      console.error('MongoDB connection error:', err);
      connectPromise = null;
      throw err;
    }
  })();
  
  return connectPromise;
}

export async function getAll() {
  await connectDB();
  return await Record.find().sort({ data: -1 }).lean();
}

export async function getById(id) {
  await connectDB();
  return await Record.findOne({ id }).lean();
}

export async function create(record) {
  await connectDB();
  const newRecord = new Record(record);
  await newRecord.save();
  return newRecord.toObject();
}

export async function update(id, record) {
  await connectDB();
  const updated = await Record.findOneAndUpdate(
    { id },
    record,
    { new: true }
  ).lean();
  return updated;
}

export async function remove(id) {
  await connectDB();
  await Record.deleteOne({ id });
}

export default Record;

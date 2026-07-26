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
  matriculaLider: String,
  matriculaEletricista: String,
  registroFoto: String,
  observacao: String,
  naoConformidade: String,
  regrasOuro: { type: Object, default: {} },
  regrasArquivos: { type: Object, default: {} }
}, { timestamps: true });

export const Record = mongoose.model('Record', recordSchema);

let connected = false;

export async function connectDB() {
  if (connected) return;
  
  try {
    const mongoUri = process.env.MONGODB_URI || 'mongodb://localhost:27017/indica';
    await mongoose.connect(mongoUri);
    connected = true;
    console.log('MongoDB connected');
  } catch (err) {
    console.error('MongoDB connection error:', err);
    throw err;
  }
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

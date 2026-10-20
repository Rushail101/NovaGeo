// Independent planning register for future purchases/projects; never posts to actual accounting.
import { supabase } from './supabase.js';

const mapItem = r => ({
  id: r.id,
  itemName: r.item_name || '',
  category: r.category || 'Other',
  estimatedAmount: Number(r.estimated_amount) || 0,
  vendorName: r.vendor_name || '',
  expectedDate: r.expected_date || '',
  status: r.status || 'Considering',
  priority: r.priority || 'Normal',
  notes: r.notes || '',
  createdAt: r.created_at || '',
  updatedAt: r.updated_at || '',
});

export async function loadPlannedExpenses() {
  const { data, error } = await supabase.from('planned_expenses').select('*').order('created_at', { ascending: false }).order('id', { ascending: false });
  if (error) return { items: [], ready: false, error: error.message || String(error) };
  return { items: (data || []).map(mapItem), ready: true, error: '' };
}

export async function savePlannedExpense(data, id = null) {
  const row = {
    item_name: String(data.itemName || '').trim(),
    category: data.category || 'Other',
    estimated_amount: Number(data.estimatedAmount) || 0,
    vendor_name: String(data.vendorName || '').trim() || null,
    expected_date: data.expectedDate || null,
    status: data.status || 'Considering',
    priority: data.priority || 'Normal',
    notes: String(data.notes || '').trim() || null,
    updated_at: new Date().toISOString(),
  };
  const query = id
    ? supabase.from('planned_expenses').update(row).eq('id', id).select().single()
    : supabase.from('planned_expenses').insert(row).select().single();
  const { data: saved, error } = await query;
  if (error) throw error;
  return mapItem(saved);
}

export async function deletePlannedExpense(id) {
  const { error } = await supabase.from('planned_expenses').delete().eq('id', id);
  if (error) throw error;
}

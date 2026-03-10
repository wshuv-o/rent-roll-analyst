export type LogType = 'system' | 'thinking' | 'grouping' | 'output' | 'flag';

export interface LogEntry {
  id: string;
  type: LogType;
  message: string;
  timestamp: Date;
  isStreaming?: boolean;
}

export interface AnonymizationMapping {
  tenantNames: Map<string, string>;
  suiteIds: Map<string, string>;
  amounts: Map<string, string>;
  reverseMap: Map<string, string>;
}

export interface ParsingInstruction {
  header_rows: number[];
  data_starts_at_row: number | null;
  column_map: {
    suite_id: string;
    tenant_name: string;
    lease_start: string;
    lease_end: string;
    gla_sqft: string;
    monthly_base_rent: string;
    base_rent_psf: string;
    recurring_charge_code: string;
    recurring_charge_amount: string;
    recurring_charge_psf: string;
    future_rent_date: string;
    future_rent_amount: string;
    future_rent_psf: string;
  };
  new_tenant_rule: string;
  skip_row_patterns: string[];
  addon_space_patterns: string[];
  confidence: 'high' | 'medium' | 'low';
  notes: string;
  // User-defined custom columns: fieldName → column letter
  custom_columns?: Record<string, string>;
}

export interface RecurringCharge {
  code: string;
  amount: number | null;
  psf: number | null;
}

export interface FutureRentIncrease {
  effective_date: string;
  monthly_amount: number | null;
  psf: number | null;
}

export interface TenantObject {
  suite_id: string;
  tenant_name: string;
  lease_start: string;
  lease_end: string;
  gla_sqft: number | null;
  monthly_base_rent: number | null;
  base_rent_psf: number | null;
  recurring_charges: RecurringCharge[];
  future_rent_increases: FutureRentIncrease[];
  notes: string;
  // User-defined custom fields: fieldName → string value
  custom_fields?: Record<string, string>;
}

// Column group definitions for visual mapping
export type ColumnGroupId = 
  | 'identity' 
  | 'lease' 
  | 'space' 
  | 'base-rent' 
  | 'charges' 
  | 'future-rent';

export interface ColumnGroup {
  id: ColumnGroupId;
  label: string;
  fields: (keyof ParsingInstruction['column_map'])[];
  fieldLabels: Record<string, string>;
}

export const COLUMN_GROUPS: ColumnGroup[] = [
  {
    id: 'identity',
    label: 'Identity',
    fields: ['suite_id', 'tenant_name'],
    fieldLabels: { suite_id: 'Suite ID', tenant_name: 'Tenant Name' },
  },
  {
    id: 'lease',
    label: 'Lease Dates',
    fields: ['lease_start', 'lease_end'],
    fieldLabels: { lease_start: 'Start', lease_end: 'End' },
  },
  {
    id: 'space',
    label: 'Space',
    fields: ['gla_sqft'],
    fieldLabels: { gla_sqft: 'GLA (SF)' },
  },
  {
    id: 'base-rent',
    label: 'Base Rent',
    fields: ['monthly_base_rent', 'base_rent_psf'],
    fieldLabels: { monthly_base_rent: 'Monthly', base_rent_psf: 'PSF' },
  },
  {
    id: 'charges',
    label: 'Recurring Charges',
    fields: ['recurring_charge_code', 'recurring_charge_amount', 'recurring_charge_psf'],
    fieldLabels: { recurring_charge_code: 'Code', recurring_charge_amount: 'Amount', recurring_charge_psf: 'PSF' },
  },
  {
    id: 'future-rent',
    label: 'Future Rent Increases',
    fields: ['future_rent_date', 'future_rent_amount', 'future_rent_psf'],
    fieldLabels: { future_rent_date: 'Date', future_rent_amount: 'Amount', future_rent_psf: 'PSF' },
  },
];

// Group span: which columns belong to a group (independent of field assignments)
export interface GroupSpan {
  groupId: ColumnGroupId;
  startCol: number;
  endCol: number;
}

// Workflow state
export type WorkflowStep = 'upload' | 'analyzing' | 'confirm' | 'parsing' | 'done';

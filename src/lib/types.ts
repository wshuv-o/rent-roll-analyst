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
}

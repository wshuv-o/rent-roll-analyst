/**
 * Rent roll type catalogue.
 *
 * Each type will eventually have its own detection policy and extraction policy.
 * For now the user selects the type manually after upload.
 */
//src/lib/rent-roll-types/index.ts
export interface RentRollType {
  id: string;
  label: string;
  description: string;
  /** true = fully wired (AI detection + extraction works) */
  implemented: boolean;
}

/** All known rent roll types */
export const RENT_ROLL_TYPES: RentRollType[] = [
  {
    id: 'regular',
    label: 'Regular Rent Roll',
    description: 'Standard layout — single header row, suite column alternates filled/empty for tenant grouping.',
    implemented: true,
  },
  {
    id: 'tenancy-schedule',
    label: 'Tenancy Schedule',
    description: 'Hierarchical blocks with "Rent Steps" / "Charge Schedules" sub-sections and repeated sub-headers per tenant.',
    implemented: true,
  },
  {
    id: 'mall-rent-roll',
    label: 'Mall Rent Roll',
    description: 'JDE EnterpriseOne mall format — multi-row tenant blocks with charge codes and future escalations.',
    implemented: true,
  },
];

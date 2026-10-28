import { sqliteTable, text, integer, check, index } from 'drizzle-orm/sqlite-core';
import { sql } from 'drizzle-orm';

export const unsubscribeBudgets = sqliteTable('unsubscribe_budgets', {
  key: text('key').primaryKey(),
  remaining: integer('remaining').notNull(),
  expires: integer('expires').notNull()
}, (table) => [check('budget_nonnegative', sql`${table.remaining} >= 0`), index('budget_expiry').on(table.expires)]);

export const unsubscribeAttempts = sqliteTable('unsubscribe_attempts', {
  key: text('key').primaryKey(),
  status: text('status').notNull(),
  reason: text('reason').notNull(),
  expires: integer('expires').notNull()
}, (table) => [index('attempt_expiry').on(table.expires)]);

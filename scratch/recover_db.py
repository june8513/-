import sqlite3
import os

def recover_database(corrupted_path, output_path):
    print(f"Attempting to recover data from {corrupted_path} to {output_path}...")
    
    try:
        # Connect to corrupted db
        src_conn = sqlite3.connect(corrupted_path)
        # Connect to new db
        dst_conn = sqlite3.connect(output_path)
        
        src_cursor = src_conn.cursor()
        dst_cursor = dst_conn.cursor()
        
        # Get all tables
        src_cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = [row[0] for row in src_cursor.fetchall()]
        
        for table in tables:
            if table.startswith('sqlite_'):
                continue
                
            print(f"Recovering table: {table}")
            try:
                # Try to fetch all data from corrupted table
                src_cursor.execute(f"SELECT * FROM {table}")
                rows = src_cursor.fetchall()
                
                if not rows:
                    continue
                    
                # Get column info for destination
                dst_cursor.execute(f"PRAGMA table_info({table})")
                cols = dst_cursor.fetchall()
                if not cols:
                    print(f"Table {table} does not exist in destination, skipping...")
                    continue
                
                col_names = [c[1] for c in cols]
                placeholders = ",".join(["?"] * len(col_names))
                
                # Insert or Ignore (to avoid duplicates from backup)
                sql = f"INSERT OR IGNORE INTO {table} ({','.join(col_names)}) VALUES ({placeholders})"
                
                dst_cursor.executemany(sql, rows)
                print(f"  Successfully recovered {len(rows)} rows for {table}")
                
            except Exception as e:
                print(f"  Failed to recover table {table}: {e}")
        
        dst_conn.commit()
        print("Recovery attempt finished.")
        
    except Exception as e:
        print(f"Recovery failed: {e}")
    finally:
        src_conn.close()
        dst_conn.close()

if __name__ == "__main__":
    corrupted = "/home/june/material-requisition/db.sqlite3.corrupted"
    current = "/home/june/material-requisition/db.sqlite3"
    
    if os.path.exists(corrupted):
        recover_database(corrupted, current)
    else:
        print("Corrupted database file not found.")

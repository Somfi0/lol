from os import path, remove, walk, rmdir
from datetime import datetime as dt, timedelta
from humanize import naturalsize
from logging import Formatter, getLogger, INFO
from logging.handlers import RotatingFileHandler
#---------------------------------------------------------------------------------------------------------------------------------------------------------#
def remove_old_files(log_path, folder, days_left):
    now = dt.now()
    removed_files_counter = 0
    megas_counter = 0
    compare_files = 0
    previous_date = dt.today() - timedelta(days=days_left)
    logger = getLogger('purge_Fab21_PRD_PL')
    logger.setLevel(INFO)
    formatter = Formatter('%(asctime)s : %(name)s : %(levelname)s : %(message)s')
    handler = RotatingFileHandler(log_path + 'purge_Fab21_PRD_PL.log', maxBytes=20000000, backupCount=1)
    handler.setLevel(INFO)
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.info(f'Start Script at: {str(now)}')
    logger.info(f'Delete Files Older Than: {str(previous_date)}')
    logger.info('-' * 136)

    try:
        for dirpath, dirnames, filenames in walk(folder):
            for file in filenames:
                if file.endswith(('.adf','.imz')):
                    full_path = path.join(dirpath, file)
                    if path.isfile(full_path):
                        try:
                            file_modify_time = dt.fromtimestamp(path.getmtime(full_path))
                        except OSError as e:
                            logger.error(f'File Modify Time Exception: {e}')
                            pass
                        if file_modify_time < previous_date:
                            try:
                                removed_files_counter += 1
                                megas_counter += path.getsize(full_path)
                                if removed_files_counter > compare_files:
                                    logger.info(f'Deleting File: {full_path} Last Modified: {str(file_modify_time)}')
                                    compare_files += 1000
                                remove(full_path)
                            except OSError as e:
                                logger.error(f'Remove File Exception: {e}')
                                pass
    except OSError as e:
        logger.error(f'Walk Exception: {e}')
        pass

    elapsed = dt.now() - now
    logger.info('-' * 136)
    logger.info(f'End Script at: {str(elapsed)}')
    return (f'Files Removed: {removed_files_counter}, Deleted: {naturalsize(megas_counter)}, Elapsed Time: {elapsed}, Runtime: {now}')
#---------------------------------------------------------------------------------------------------------------------------------------------------------#
def Log_History(log_path, num_of_removed_files):
    logger = getLogger('purge_Fab21_PRD_PL_History')
    logger.setLevel(INFO)
    formatter = Formatter('%(asctime)s : %(name)s : %(levelname)s : %(message)s')
    handler = RotatingFileHandler(log_path + 'purge_Fab21_PRD_PL_History.log', maxBytes=20000000, backupCount=1)
    handler.setLevel(INFO)
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.info(f'{num_of_removed_files}')
    logger.info('-' * 136)
#---------------------------------------------------------------------------------------------------------------------------------------------------------#
if __name__ == "__main__":
   log_path = 'C:\\Scripts\\CFFAB\\Logs\\'
   Path_To_Purge = '\\\\aligntech.com\\PL2AFAB\\pl2_fabdata\\fab21'
   num_of_removed_files = remove_old_files(log_path, Path_To_Purge, 30)
   Log_History(log_path, num_of_removed_files)

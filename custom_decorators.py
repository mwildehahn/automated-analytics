from emailfunctions import send_email
import traceback

def logerrors(to_email, cron_file,
    subject='Cron Error', include_traceback=True):
    """ Decrator to send an email if a function fails.

    args:
        to_email: needs to be a list of emails to notify, ie.
            ['mwhahn@gmail.com']
        cron_file: some string to include in the email noting the file that
            you're running, this will also be contained in the traceback but
            this is also helpful.
    kwargs:
        subject: subject of email
        include_traceback: whether or not to include traceback in email

    example:

    @logerrors(['mwhahn@gmail.com'], 'Sales Daily Report')
    def main():
        # all the functions to run your script

    if __name__ == '__main__':
        main()
    """

    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception, e:
                message = 'Error runinng cron file: %s.\n\n' % cron_file
                if include_traceback:
                    message += traceback.format_exc()
                send_email(to=to_email, subject=subject, message=message)
        return wrapper
    return decorator
